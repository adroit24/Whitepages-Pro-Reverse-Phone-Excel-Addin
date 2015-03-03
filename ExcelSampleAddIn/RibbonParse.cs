using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Utilities;
using Newtonsoft.Json;
using WebService;
using System.IO;
using System.Collections.Specialized;


namespace ExcelSampleAddIn
{
    public partial class RibbonParse
    {
        private void RibbonParse_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void WPData_Button_Click(object sender, RibbonControlEventArgs e)
        {

            //Get active sheet
            Worksheet activeWorksheet = ((Worksheet)Globals.ThisAddIn.Application.ActiveSheet);

            //Get all the phone numbers
            //var cellValue = (string)(activeWorksheet.Cells[1, 1] as Range).Value;
            List<string> phoneList = new List<string>();
            string apiKey = "";

            Range usedRange = activeWorksheet.UsedRange;

            //Get API Key
            apiKey = Convert.ToString((activeWorksheet.Cells[1, 1] as Range).Value);

            for (int rCnt = 2; rCnt <= usedRange.Rows.Count; rCnt++)
            {

                phoneList.Add(Convert.ToString((usedRange.Cells[rCnt, 1] as Range).Value));

            }

            //Get Whitepages data
            PopulateWhitepagesData(phoneList, apiKey);

        }


        void PopulateWhitepagesData(List<string> phoneList, string apiKey)
        {
            int statusCode = -1;
            string description = string.Empty;
            string errorMessage = string.Empty;

            int rowNumToPopulate = 2;

            for (int i = 0; i < phoneList.Count; i++)
            {
                
                NameValueCollection nameValues = new NameValueCollection();

                nameValues["phone"] = phoneList[i];
                nameValues["api_key"] = apiKey;
                  
                WhitePagesWebService webService = new WhitePagesWebService();
                // Call method ExecuteWebRequest to execute backend API and return response stream.
                Stream responseStream = webService.ExecuteWebRequest(nameValues, ref statusCode, ref description, ref errorMessage);

                // Checking respnseStream null and status code.
                if (statusCode == 200 && responseStream != null)
                {
                    // Reading response stream to StreamReader.
                    StreamReader reader = new StreamReader(responseStream);

                    // Convert stream reader to string JSON.
                    string responseInJson = reader.ReadToEnd();

                    // Dispose response stream
                    responseStream.Dispose();

                    // Calling ParsePhoneLookupResult to parse the response JSON in data Result class.
                    Result resultData = ParsePhoneLookupResult(responseInJson);

                    PopulateExcel(resultData, rowNumToPopulate++);

                }
            }
        }

        private void PopulateExcel(Result resultData, int rowNumToPopulate)
        {
            PopulateTopRowHeaders();

            //Get active sheet
            Worksheet activeWorksheet = ((Worksheet)Globals.ThisAddIn.Application.ActiveSheet);

            string rowNum = Convert.ToString(rowNumToPopulate);
            
            Range numColumn = activeWorksheet.get_Range("B"+rowNum);
            numColumn.Value2 = resultData.Phone.PhoneNumber;

            Range isValidColumn = activeWorksheet.get_Range("C"+rowNum);
            isValidColumn.Value2 = resultData.Phone.IsValid;

            Range countryColumn = activeWorksheet.get_Range("D"+rowNum);
            countryColumn.Value2 = resultData.Phone.CountryCallingCode;

            Range isPrepaidcolumn = activeWorksheet.get_Range("E"+rowNum);
            isPrepaidcolumn.Value2 = resultData.Phone.IsPrepaid;

            Range lineTypeColumn = activeWorksheet.get_Range("F"+rowNum);
            lineTypeColumn.Value2 = resultData.Phone.PhoneType;

            Range carrierColumn  = activeWorksheet.get_Range("G"+rowNum);
            carrierColumn.Value2 = resultData.Phone.Carrier;

            Range dncColumn = activeWorksheet.get_Range("H" + rowNum);
            dncColumn.Value2 = resultData.Phone.DndStatus;

            //Person Data
            if (resultData.GetPeople() != null && (resultData.GetPeople().Length > 0))
            {
                Range personTypeColumn = activeWorksheet.get_Range("I" + rowNum);
                personTypeColumn.Value2 = resultData.GetPeople()[0].PersonType;

                Range nameColumn = activeWorksheet.get_Range("J" + rowNum);
                nameColumn.Value2 = resultData.GetPeople()[0].PersonName;

                Range genderColumn = activeWorksheet.get_Range("K" + rowNum);
                genderColumn.Value2 = resultData.GetPeople()[0].Gender;

                Range ageColumn = activeWorksheet.get_Range("L" + rowNum);
                ageColumn.Value2 = resultData.GetPeople()[0].AgeRange;
            }
            
            
            //Location Data
            if (resultData.GetLocation() != null && (resultData.GetLocation().Length > 0))
            {
                Range streetAddress1Column = activeWorksheet.get_Range("M" + rowNum);
                streetAddress1Column.Value2 = resultData.GetLocation()[0].StandardAddressLine1;

                Range streetAddress2Column = activeWorksheet.get_Range("N" + rowNum);
                streetAddress2Column.Value2 = resultData.GetLocation()[0].StandardAddressLine2;

                Range locationColumn = activeWorksheet.get_Range("O" + rowNum);
                locationColumn.Value2 = resultData.GetLocation()[0].StandardAddressLocation;

                Range isMailReceivableColumn = activeWorksheet.get_Range("P" + rowNum);
                isMailReceivableColumn.Value2 = resultData.GetLocation()[0].ReceivingMail;

                Range deliveryPointColumn = activeWorksheet.get_Range("Q" + rowNum);
                deliveryPointColumn.Value2 = resultData.GetLocation()[0].DeliveryPoint;
            }

        }

        private void PopulateTopRowHeaders()
        {
            //Get active sheet
            Worksheet activeWorksheet = ((Worksheet)Globals.ThisAddIn.Application.ActiveSheet);

            string rowNum = Convert.ToString(1);

            Range numColumn = activeWorksheet.get_Range("B" + rowNum);
            numColumn.Value2 = "Phone Number";

            Range isValidColumn = activeWorksheet.get_Range("C" + rowNum);
            isValidColumn.Value2 = "Is Valid";

            Range countryColumn = activeWorksheet.get_Range("D" + rowNum);
            countryColumn.Value2 = "Country Calling Code";

            Range isPrepaidcolumn = activeWorksheet.get_Range("E" + rowNum);
            isPrepaidcolumn.Value2 = "Is Prepaid";

            Range lineTypeColumn = activeWorksheet.get_Range("F" + rowNum);
            lineTypeColumn.Value2 = "Line Type";

            Range carrierColumn = activeWorksheet.get_Range("G" + rowNum);
            carrierColumn.Value2 = "Carrier";

            Range dncColumn = activeWorksheet.get_Range("H" + rowNum);
            dncColumn.Value2 = "Do not Call Registered";

            //Person Data
            Range personTypeColumn = activeWorksheet.get_Range("I" + rowNum);
            personTypeColumn.Value2 = "Person or Business";

            Range nameColumn = activeWorksheet.get_Range("J" + rowNum);
            nameColumn.Value2 = "Name";

            Range genderColumn = activeWorksheet.get_Range("K" + rowNum);
            genderColumn.Value2 = "Gender";

            Range ageColumn = activeWorksheet.get_Range("L" + rowNum);
            ageColumn.Value2 = "Age Range";

            //Location Data
            Range streetAddress1Column = activeWorksheet.get_Range("M" + rowNum);
            streetAddress1Column.Value2 = "Street Address Line 1";

            Range streetAddress2Column = activeWorksheet.get_Range("N" + rowNum);
            streetAddress2Column.Value2 = "Street Address Line 2";

            Range locationColumn = activeWorksheet.get_Range("O" + rowNum);
            locationColumn.Value2 = "Location";

            Range isMailReceivableColumn = activeWorksheet.get_Range("P" + rowNum);
            isMailReceivableColumn.Value2 = "ReceivingMail";

            Range deliveryPointColumn = activeWorksheet.get_Range("Q" + rowNum);
            deliveryPointColumn.Value2 = "DeliveryPoint";
        }


        /// <summary>
        /// This method parse the Phone Lookup data to class Result.
        /// </summary>
        /// <param name="responseInJson">responseInJson</param>
        /// <returns>Result</returns>
        private Result ParsePhoneLookupResult(string responseInJson)
        {
            // Creating PhoneLookupData object to fill the phone lookup data.
            Result resultData = new Result();

            try
            {
                // responseInJson to DeserializeObject
                dynamic jsonObject = JsonConvert.DeserializeObject(responseInJson);

                if (jsonObject != null)
                {
                    // Take the dictionary object from jsonObject.
                    dynamic dictionaryObj = jsonObject.dictionary;

                    if (dictionaryObj != null)
                    {
                        string phoneKey = string.Empty;

                        // Take the phone key from result node of jsonObject
                        foreach (var data in jsonObject.results)
                        {
                            phoneKey = data.Value;
                            break;
                        }

                        #region Phone Data

                        // Checking phone key null or empty.
                        if (!string.IsNullOrEmpty(phoneKey))
                        {
                            // Get phone key object from dictionaryObj using phoneKey.
                            dynamic phoneKeyObject = dictionaryObj[phoneKey];

                            if (phoneKeyObject != null)
                            {
                                // Creating phoneData object to fill the phone lookup data.
                                Phone phoneData = new Phone();

                                // Extracting lineType,phoneNumber, countryCallingCode, carrier, doNotCall status, spamScore from phoneKeyObject.

                                if (phoneKeyObject["is_valid"] != null)
                                {
                                    phoneData.IsValid = (bool)phoneKeyObject["is_valid"];

                                }

                                if (phoneKeyObject["line_type"] != null)
                                {
                                    phoneData.PhoneType = (string)phoneKeyObject["line_type"];

                                }

                                if (phoneKeyObject["phone_number"] != null)
                                {
                                    phoneData.PhoneNumber = (string)phoneKeyObject["phone_number"];

                                    //Increment counter here for Phone
                                    resultData.DataCounters++;
                                }

                                if (phoneKeyObject["country_calling_code"] != null)
                                {
                                    phoneData.CountryCallingCode = (string)phoneKeyObject["country_calling_code"];

                                }

                                if (phoneKeyObject["carrier"] != null)
                                {
                                    phoneData.Carrier = (string)phoneKeyObject["carrier"];

                                }
                                if (phoneKeyObject["do_not_call"] != null)
                                {
                                    phoneData.DndStatus = (bool)(phoneKeyObject["do_not_call"]);

                                }


                                dynamic spamScoreObj = phoneKeyObject.reputation;
                                if (spamScoreObj != null)
                                {
                                    phoneData.SpamScore = (string)spamScoreObj["spam_score"];

                                }

                                if (phoneKeyObject["is_prepaid"] != null)
                                {
                                    phoneData.IsPrepaid = (bool)(phoneKeyObject["is_prepaid"]);

                                }

                                resultData.Phone = phoneData;

                        #endregion

                                #region Person and Business
                                // Starting to extarct the person information.
                                dynamic phoneKeyObjectBelongsToObj = phoneKeyObject.belongs_to;

                                List<string> personKeyListFromBelongsTo = new List<string>();
                                List<string> locationKeyList = new List<string>();


                                // Lets get the basic location from associated location of phone.
                                string phoneAssociatedLocation = "";
                                // Extracting location key from Phone details under best_location object.
                                dynamic bestLocationFromPhoneObj = phoneKeyObject.best_location;
                                if (bestLocationFromPhoneObj != null)
                                {
                                    dynamic bestLocationIdFromPhoneObj = bestLocationFromPhoneObj.id;
                                    if (bestLocationIdFromPhoneObj != null)
                                    {
                                        phoneAssociatedLocation = ((string)bestLocationIdFromPhoneObj["key"]);
                                    }
                                }

                            
                                if (phoneKeyObjectBelongsToObj != null)
                                {

                                    // Creating list of person key from phoneKeyObjectBelongsToObj.
                                    foreach (var data in phoneKeyObjectBelongsToObj)
                                    {
                                        dynamic belongsToObj = data.id;
                                        if (belongsToObj != null)
                                        {
                                            string personKeyFromBelongsTo = (string)belongsToObj["key"];

                                            if (!string.IsNullOrEmpty(personKeyFromBelongsTo))
                                            {
                                                personKeyListFromBelongsTo.Add(personKeyFromBelongsTo);
                                            }

                                        }
                                    }
                                }

                                List<People> peopleList = new List<People>();

                                if (personKeyListFromBelongsTo.Count > 0)
                                {
                                    //Increment counter for Person
                                    resultData.DataCounters++;

                                    People people = null;
                                    dynamic personKeyObject = null;

                                    foreach (string personKey in personKeyListFromBelongsTo)
                                    {
                                        people = new People();
                                        personKeyObject = dictionaryObj[personKey];
                                        if (personKeyObject != null)
                                        {
                                            people = ParsePersonData(personKeyObject);
                                #endregion

                                            #region Location
                                            // Collecting Locations Key. if best_location node exist other wise will take location key from locations node
                                            string locationKey = string.Empty;
                                            if (personKeyObject["best_location"] != null)
                                            {

                                                dynamic personBestLocationObj = personKeyObject.best_location;
                                                if (personBestLocationObj != null)
                                                {
                                                    //Empty the Phone associated location we added earlier
                                                    dynamic bestLocationIdObj = personBestLocationObj.id;
                                                    if (bestLocationIdObj != null)
                                                    {
                                                        locationKey = (string)bestLocationIdObj["key"];
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                if (personKeyObject["locations"] != null)
                                                {

                                                    dynamic locationsPerPersonObj = personKeyObject.locations;
                                                    if (locationsPerPersonObj != null)
                                                    {
                                                        foreach (var personLocation in locationsPerPersonObj)
                                                        {
                                                            dynamic locationIdObj = personLocation.id;
                                                            if (locationIdObj != null)
                                                            {
                                                                locationKey = (string)locationIdObj["key"];
                                                                break;
                                                            }
                                                        }
                                                    }
                                                }
                                            }

                                            locationKeyList.Add(locationKey);
                                            peopleList.Add(people);
                                        }
                                    }

                                    resultData.SetPeople(peopleList.ToArray());
                                    List<Location> locationList = ParseLocationData(dictionaryObj, locationKeyList);
                                    if (locationList.Count > 0)
                                    {
                                        //Increment counter for location
                                        resultData.DataCounters++;
                                        resultData.SetLocation(locationList.ToArray());
                                    }
                                }

                                if (resultData.GetLocation() == null && phoneAssociatedLocation.Length > 0)
                                {
                                    locationKeyList.Add(phoneAssociatedLocation);
                                    List<Location> locationList = ParseLocationData(dictionaryObj, locationKeyList);

                                    resultData.SetLocation(locationList.ToArray());
                                }
                                            #endregion
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                
            }

            return resultData;
        }

        private static People ParsePersonData(dynamic personKeyObject)
        {

            People people = new People();

            // Get phoneKeyIdObj from personKeyObject.
            dynamic phoneKeyIdObj = personKeyObject.id;

            if (phoneKeyIdObj != null)
            {
                // Get person type from phoneKeyIdObj.
                people.PersonType = phoneKeyIdObj["type"];
            }

            // phoneKeyNamesObj from name node of personKeyObject.
            dynamic phoneKeyNamesObj = personKeyObject.names;

            if (phoneKeyNamesObj != null)
            {
                string firstName = string.Empty;
                string lastName = string.Empty;
                string middleName = string.Empty;

                foreach (var name in phoneKeyNamesObj)
                {
                    firstName = (string)name["first_name"];
                    middleName = (string)name["middle_name"];
                    lastName = (string)name["last_name"];
                }

                people.PersonName = firstName + " " + lastName;
            }
            else
            {
                people.PersonName = personKeyObject.name;
            }

            if (personKeyObject.gender != null)
            {
                people.Gender = personKeyObject.gender;
            }
            else
            {
                people.Gender = "Unknown";
            }

            dynamic phoneKeyAgeObj = personKeyObject.age_range;
            if (phoneKeyAgeObj != null)
            {
                people.AgeRange += phoneKeyAgeObj["start"];
                people.AgeRange += "-";
                people.AgeRange += phoneKeyAgeObj["end"];
            }
            else
            {
                people.AgeRange += "Unknown";
            }

            return people;
        }

        private static List<Location> ParseLocationData(dynamic dictionaryObj, List<string> locationKeyList)
        {
            List<Location> locationList = new List<Location>();
            Location location = null;

            // Extracting all location for all locationKeyList from locationKeyObject.
            foreach (string locationKey in locationKeyList)
            {
                location = new Location();

                dynamic locationKeyObject = dictionaryObj[locationKey];

                if (locationKeyObject != null)
                {
                    location.StandardAddressLine1 = (string)locationKeyObject["standard_address_line1"];
                    location.StandardAddressLine2 = (string)locationKeyObject["standard_address_line2"];
                    location.StandardAddressLocation = (string)locationKeyObject["standard_address_location"];
                    if (locationKeyObject["is_receiving_mail"] != null)
                    {
                        location.ReceivingMail = (bool)(locationKeyObject["is_receiving_mail"]);
                    }

                    if ((string)locationKeyObject["usage"] != null)
                    {
                        location.Usage = (string)locationKeyObject["usage"];
                    }
                    else
                    {
                        location.Usage = "Unknown";
                    }
                    if ((string)locationKeyObject["delivery_point"] != null)
                    {
                        location.DeliveryPoint = (string)locationKeyObject["delivery_point"];
                    }
                    else
                    {
                        location.DeliveryPoint = "Unknown";
                    }
                    locationList.Add(location);
                }
            }
            return locationList;
        }

    }
}
