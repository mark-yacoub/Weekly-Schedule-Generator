/*
 * Create a list of 2 deacons groups weekly based on their language of perference
 * the list is read from an excel sheet created through Google forms
 * the list has the name, age and language preference.
 * each deacons group should consist of members of each age group
*/

using System;
using System.Collections.Generic;
using System.Linq;
using ExcelLibrary.Office.Excel;

namespace AltarSchedule
{
    class Program
    {
        static int NUMBER_OF_SERVANTS_IN_THE_ALTAR;

        enum Language
        {
            French, English, EnglishAndFrench
        }

        struct DeaconData
        {
            public string Name;
            public double Age;
            public Language Language;
        }


        static void Main()
        {
            // get the number of the desired number of deacons in each group
            Console.WriteLine("Enter the number of altar deacons");
            string userInput = Console.ReadLine();
            NUMBER_OF_SERVANTS_IN_THE_ALTAR = Convert.ToInt32(userInput);

            Worksheet sheet = OpenSheet();
            List<DeaconData> deacons = PopulateDeaconsList(sheet);

            List<DeaconData> frenchList = GetLanguageGroups(deacons, Language.French);
            List<DeaconData> englishList = GetLanguageGroups(deacons, Language.English);

            List<List<DeaconData>> frenchAgeGroups = CreateAltarGroups(frenchList);
            List<List<DeaconData>> englishAgeGroups = CreateAltarGroups(englishList);

            IEnumerable<List<DeaconData>> bothLists = frenchAgeGroups.Concat(englishAgeGroups);
            // the maximum number of deacons is the largest number of people within the same age group in any language group
            int maximumNumberOfDeaconsInList = bothLists.Max(e => e.Count);
            
            CreateScheduleSheet(frenchAgeGroups, englishAgeGroups, maximumNumberOfDeaconsInList);
        }

        // start readong the excel sheet
        private static Worksheet OpenSheet()
        {
            try
            {
                Console.WriteLine("Reading Excel File...");
                Workbook book = Workbook.Open("St George Deacons Sign Up Sheet (Responses).xls");
                return book.Worksheets[0];
            }
            catch (Exception ex)
            {
                if (ex.ToString().Contains("used"))
                    Console.WriteLine("most probably the file is open, please close it; otherwise");
                Console.WriteLine("Failed to Open file \n make sure the file is saved as XLS (old Excel Format) and is on the Desktop");
                // hold for the user to know something has going wrong.
                while (true) ;
            }
        }

        // Create a list of all deacons
        private static List<DeaconData> PopulateDeaconsList(Worksheet sheet)
        {
            Console.WriteLine("Creating the List of all deacons...");

            List<DeaconData> deaconsList = new List<DeaconData>();
            // format is: (1)Name   \t  (2)Age  \t  (3)Language
            for (int rowIndex = 1; rowIndex <= sheet.Cells.LastRowIndex; rowIndex++)
            {
                Row row = sheet.Cells.GetRow(rowIndex);
                DeaconData dData = new DeaconData();

                dData.Name = row.GetCell(1).StringValue;
                dData.Age = Convert.ToDouble(row.GetCell(2).StringValue);

                string language = row.GetCell(3).StringValue;
                if (language == "French")
                    dData.Language = Language.French;
                else if (language == "English")
                    dData.Language = Language.English;
                else
                    dData.Language = Language.EnglishAndFrench;

                deaconsList.Add(dData);
            }
            return deaconsList;
        }

        // sort them in 2 lists based on their languages to create separate language groups
        // the function is called for every language that needs to be created
        private static List<DeaconData> GetLanguageGroups(List<DeaconData> altarGroups, Language listLanguage)
        {
            Console.WriteLine("Sorting the deacons based on their preferred languages...");

            List<DeaconData> languageAgeGroups = new List<DeaconData>();
            foreach (DeaconData deacon in altarGroups)
            {
                if (deacon.Language == listLanguage || deacon.Language == Language.EnglishAndFrench)
                    languageAgeGroups.Add(deacon);
            }
            return languageAgeGroups;
        }

        private static List<List<DeaconData>> CreateAltarGroups(List<DeaconData> deaconsList)
        {
            Console.WriteLine("Splitting them into age groups...");

            // sort the deacons in the list by their ages
            List<DeaconData> sortedList = deaconsList.OrderBy(x => x.Age).ToList();

            List<List<DeaconData>> ageGroups = new List<List<DeaconData>>();
            int listIndex = 0;
            // The number of different age groups will be based on the number of deacons assigned in the altar
            // that gives each age group an equal change of having a turn, with a good age distribution
            for (int i = 0; i < NUMBER_OF_SERVANTS_IN_THE_ALTAR; i++)
            {
                int range = deaconsList.Count / NUMBER_OF_SERVANTS_IN_THE_ALTAR;
                // the first iteration is expected to separate the younger age groups
                // if the number of deacons cannot be separated evenly across #NUMBER_OF_SERVANTS_IN_THE_ALTAR
                // the kids should have the largest number of people in their group, giving more preference to the elders
                if (i == 0)
                    range += deaconsList.Count % NUMBER_OF_SERVANTS_IN_THE_ALTAR;

                // take only a part of the full list within the specified range and assign it to separate new list
                List<DeaconData> ageGroup = sortedList.GetRange(listIndex, range);
                ageGroups.Add(ageGroup);

                listIndex += range;
            }
            return ageGroups;
        }

        // Write to an Excel Sheet the final schedule
        private static void CreateScheduleSheet(List<List<DeaconData>> frenchGroups, List<List<DeaconData>> englishGroups, int rotationCount)
        {
            Console.WriteLine("Finally Creating the Schedule...");

            string fileName = "AltarServiceSchedule.xls";
            Workbook workbook = new Workbook();
            Worksheet worksheet = new Worksheet("Altar 2015");
            worksheet.Cells[0, 0] = new Cell("French Liturgy");
            worksheet.Cells[0, NUMBER_OF_SERVANTS_IN_THE_ALTAR + 2] = new Cell("English Liturgy");

            // use random numbers to randonly assign people
            // create pools of used generated random numbers, one for each group to avoid repitition 
            List<List<int>> usedRandomNumbers = new List<List<int>>();
            for (int i = 0; i < NUMBER_OF_SERVANTS_IN_THE_ALTAR*2; i++)
                // that insures the same person is not assigned twice on the same day
                usedRandomNumbers.Add(new List<int>());

            Random rand = new Random();

            // the rotation count ensures that every single person has gone at least once
            for (int i = 1; i <= rotationCount; i++)
            {
                List<string> servantsOfTheDay = new List<string>();
                for (int j = 0; j < NUMBER_OF_SERVANTS_IN_THE_ALTAR; j++)
                {
                    int frenchIndex;
                    int englishIndex;

                    do
                    {
                        do
                        {
                            // find a random number within the rage of the size of the group
                            frenchIndex = rand.Next(frenchGroups[j].Count);
                            // check if the random numbers used are equal of size to the size if the language group
                            // that would mean everyone has gone once
                            if (usedRandomNumbers[j*2].Count == frenchGroups[j].Count)
                                // clear the random number pool to allow for new repitition
                                usedRandomNumbers[j*2].Clear();
                        // if the random has been found used, loop again for new number
                        } while (usedRandomNumbers[j*2].Contains(frenchIndex));
                        usedRandomNumbers[j*2].Add(frenchIndex);

                        do
                        {
                            englishIndex = rand.Next(englishGroups[j].Count);
                            if (usedRandomNumbers[j*2 + 1].Count == englishGroups[j].Count)
                                usedRandomNumbers[j*2 + 1].Clear();
                        } while (usedRandomNumbers[j*2 + 1].Contains(englishIndex));
                        usedRandomNumbers[j*2 + 1].Add(englishIndex);

                    // ensure that the same person is not assigned twice on the same day
                    // might happen to people with 2 languages preferences
                    } while (servantsOfTheDay.Contains(englishGroups[j][englishIndex].Name) ||
                             servantsOfTheDay.Contains(frenchGroups[j][frenchIndex].Name));
                    servantsOfTheDay.Add(englishGroups[j][englishIndex].Name);
                    servantsOfTheDay.Add(frenchGroups[j][frenchIndex].Name);

                    worksheet.Cells[i, j] = new Cell(frenchGroups[j][frenchIndex].Name);
                    worksheet.Cells[i, j + NUMBER_OF_SERVANTS_IN_THE_ALTAR + 2] = new Cell(englishGroups[j][englishIndex].Name);

                    servantsOfTheDay.Clear();
                }
            }

            // A bug in the library that corrupts the file if the size is small
            // the fix: add new empty cells
            int col, row = rotationCount;
            for (int i = 0; i < 150; i++)
            {
                col = 1;
                for (int j = 0; j < 10; j++)
                {
                    worksheet.Cells[row + 5, col] = new Cell(" ");
                    col++;
                }
                row++;
            }

            workbook.Worksheets.Add(worksheet);
            workbook.Save(fileName);
            // start the file for the user
            System.Diagnostics.Process.Start(fileName);
        }

    }
}