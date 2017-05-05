using System;
using System.IO;
using System.Xml;
using System.Linq;
using System.Collections;
using System.Configuration;
using System.Globalization;
using System.Collections.Generic;

namespace XmlParserSample
{
    public class LoadAndManipulateXml
    {
        #region data members

        private static readonly string SourceFolder = ConfigurationManager.AppSettings["SourceFolder"];

        private static readonly string TargetFolder = ConfigurationManager.AppSettings["TargetFolder"];

        private static readonly Random Random = new Random();

        private const string Pattern = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";

        private const string CountrySettingsInputXml = "<CountrySettingsTable>" +
                                         "<CountrySettingsReference>United States - Input</CountrySettingsReference>" +
                                         "<AreaMeasurementUnit>Square Feet</AreaMeasurementUnit>" +
                                         "<AreaMeasurementAbbreviation>SF</AreaMeasurementAbbreviation>" +
                                         "<CurrencyName>Dollars</CurrencyName>" +
                                         "<CurrencySymbol>$</CurrencySymbol>" +
                                         "<SymbolPosition>Immediately Before Currency</SymbolPosition>" +
                                         "<DecimalSymbol>Period</DecimalSymbol>" +
                                         "<ThousandsSeparator>Comma</ThousandsSeparator>" +
                                         "<ConversionRate BaseCurrency=\"Dollars\">1</ConversionRate>" +
                                         "<DateFormat>MonthDayYear</DateFormat>" +
                                         "</CountrySettingsTable>";

        private const string CountrySettingsOutputXml = "<CountrySettingsTable>" +
                                       "<CountrySettingsReference>United States - Output</CountrySettingsReference>" +
                                       "<AreaMeasurementUnit>Square Feet</AreaMeasurementUnit>" +
                                       "<AreaMeasurementAbbreviation>SqFt</AreaMeasurementAbbreviation>" +
                                       "<CurrencyName>Dollars</CurrencyName>" +
                                       "<CurrencySymbol>$</CurrencySymbol>" +
                                       "<SymbolPosition>Immediately Before Currency</SymbolPosition>" +
                                       "<DecimalSymbol>Period</DecimalSymbol>" +
                                       "<ThousandsSeparator>Comma</ThousandsSeparator>" +
                                       "<ConversionRate BaseCurrency=\"Dollars\">1</ConversionRate>" +
                                       "<DateFormat>MonthDayYear</DateFormat>" +
                                       "</CountrySettingsTable>";

        private static string _logFileName = string.Empty;

        private static XmlDocument _xmlDocument;

        #endregion

        public static void LoadXmlFilesAndUpdate()
        {
            var dirInfo = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory + SourceFolder);

            if (!dirInfo.Exists)
            {
                Console.WriteLine("Source directory (\"" + SourceFolder + "\") not found.\n");
                return;
            }

            Console.WriteLine("Getting all files from \"" + SourceFolder + "\" folder.\n");

            var files = dirInfo.GetFiles("*.xml");

            if (!files.Any())
            {
                Console.WriteLine("No any file found in \"" + SourceFolder + "\" folder.\n");
                return;
            }
            Console.WriteLine("Total " + files.Count() + " files found in \"" + SourceFolder + "\" location.\n");

            _xmlDocument = new XmlDocument();

            foreach (var file in files)
            {
                _logFileName = file.Name.Replace(".xml", "");

                Console.WriteLine("Processing \"" + file.Name + "\" file...\n");
                try
                {
                    _xmlDocument.Load(file.FullName);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error in parsing xml file.\nError: " + ex.Message + "\n");
                    continue;
                }

                var xmlNodeListLeases = _xmlDocument.SelectNodes("/reXML/InputAssumptions/LeaseData/Leases/Lease");

                if (xmlNodeListLeases == null || xmlNodeListLeases.Count == 0)
                {
                    WriteLog("<Lease> node not found. Attempted path- /reXML/InputAssumptions/LeaseData/Leases/Lease \n");
                }
                else
                {
                    UpdateLeaseNodes(xmlNodeListLeases);
                }

                var xmlNodeProperty = _xmlDocument.SelectSingleNode("/reXML/InputAssumptions/PropertyData");

                if (xmlNodeProperty == null)
                {
                    WriteLog("<PropertyData> node not found. Attempted path- /reXML/InputAssumptions/PropertyData \n");
                }
                else
                {
                    UpdatePropertyData(xmlNodeProperty);

                    UpdateMakretLeasingAssumptions(xmlNodeProperty);

                    var xmlNodePropertyReference = _xmlDocument.SelectSingleNode("/reXML/PropertyReference");

                    if (xmlNodePropertyReference != null)
                    {
                        ApplyMarketLeasingAssumptionsUpdates(xmlNodePropertyReference.InnerText);
                    }
                }

                try
                {
                    dirInfo = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory + TargetFolder);
                    if (!dirInfo.Exists) dirInfo.Create();
                    var fileInfo = new FileInfo(dirInfo.FullName + "\\" + _logFileName + "_PROC.xml");
                    if (file.Exists) fileInfo.Delete();
                    _xmlDocument.Save(fileInfo.FullName);
                }
                catch (Exception ex)
                {
                    WriteLog("Error while saving file to \"" + TargetFolder + "\" folder. Error: " + ex.Message);
                    Console.WriteLine("Error while saving file to \"" + TargetFolder + "\" folder. Error: " + ex.Message);
                }
            }
        }

        private static void UpdateLeaseNodes(IEnumerable leaseNodes)
        {
            var stringList = new List<string>();

            foreach (XmlNode leaseNode in leaseNodes)
            {
                #region if lease id is duplicate

                if (leaseNode.Attributes != null && leaseNode.Attributes.Count != 0 && leaseNode.Attributes["LeaseReference"] != null)
                {
                    var xmlAttributeLeaseReference = leaseNode.Attributes["LeaseReference"];

                    if (!string.IsNullOrEmpty(xmlAttributeLeaseReference.Value))
                    {
                        if (stringList.Contains(xmlAttributeLeaseReference.Value))
                        {
                            xmlAttributeLeaseReference.Value += GetRandomString();

                            var leaseReferenceNode = leaseNode.SelectSingleNode("LeaseID/LeaseReference");

                            if (leaseReferenceNode == null)
                            {
                                WriteLog("<LeaseReference> node not found. Searched in: " + leaseNode.InnerXml);
                            }
                            else
                            {
                                leaseReferenceNode.InnerText = xmlAttributeLeaseReference.Value;
                            }
                        }
                        stringList.Add(xmlAttributeLeaseReference.Value);
                    }
                }

                #endregion

                var xmlNodeListFreeRents = leaseNode.SelectNodes("FreeRents/FreeRent");

                if (xmlNodeListFreeRents != null && xmlNodeListFreeRents.Count != 0)
                {
                    UpdateFreeRentAmount(xmlNodeListFreeRents);
                }

                var tenantStatusNode = leaseNode.SelectSingleNode("TenantStatus");

                if (tenantStatusNode == null)
                {
                    WriteLog("<TenantStatus> node not found. Searched in: " + leaseNode.InnerXml);
                    continue;
                }

                if (tenantStatusNode.Attributes == null || tenantStatusNode.Attributes["SubStatus"] == null)
                {
                    WriteLog("\"SubStatus\" attribute not found in <TenantStatus> node. Searched in: " + leaseNode.InnerXml);
                    continue;
                }

                var xmlAttributeSubStatus = tenantStatusNode.Attributes["SubStatus"].Value;

                if (string.IsNullOrEmpty(xmlAttributeSubStatus))
                {
                    continue;
                }

                #region if tenant status is M/M

                if (xmlAttributeSubStatus.ToUpper().Equals(@"M/M"))
                {

                    var startDateNode = leaseNode.SelectSingleNode("StartDate");

                    if (startDateNode == null)
                    {
                        WriteLog("<StartDate> node not found or RelativeTo attribute is null. Searched in: " +
                                 tenantStatusNode.InnerXml);
                        continue;
                    }

                    if (startDateNode.Attributes == null || startDateNode.Attributes["RelativeTo"] == null)
                    {
                        WriteLog("\"RelativeTo\" attribute not found in <StartDate> node. Searched in: " +
                                 tenantStatusNode.InnerXml);
                        continue;
                    }
                    var xmlAttributeRelativeTo = startDateNode.Attributes["RelativeTo"].Value;

                    if (xmlAttributeRelativeTo.Equals("ProjectionStart"))
                    {
                        WriteLog("RelativeTo attribute is \"ProjectionStart\" Searched in: " + startDateNode.InnerXml);
                        continue;
                    }

                    if (!xmlAttributeRelativeTo.Equals("Absolute"))
                    {
                        continue;
                    }

                    DateTime startDate;

                    if (
                        !DateTime.TryParseExact(startDateNode.InnerText, "yyyy-MM-dd", null, DateTimeStyles.None,
                            out startDate))
                    {
                        continue;
                    }
                    var newTerm = (int)Math.Truncate(DateTime.Now.Date.Subtract(startDate).TotalDays / 30) + 13;

                    var termMonthsNode = leaseNode.SelectSingleNode("TermMonths");

                    if (termMonthsNode == null)
                    {
                        WriteLog("<TermMonths> node not found. Searched in: " + startDateNode.InnerXml);
                        continue;
                    }
                    termMonthsNode.InnerText = newTerm.ToString(CultureInfo.InvariantCulture);
                }

                #endregion

                #region if tenant status is contarct

                else
                {
                    var xmlNodeBaseRentEntry = leaseNode.SelectNodes("RentalIncome/BaseRent/BaseRentEntry");
                    if (xmlNodeBaseRentEntry == null || xmlNodeBaseRentEntry.Count == 0)
                    {
                        WriteLog("<BaseRentEntry> node not found. Searched in: " + leaseNode.InnerXml);
                        continue;
                    }

                    UpdateBaseRentEntry(xmlNodeBaseRentEntry, xmlAttributeSubStatus);
                }

                #endregion
            }
        }

        private static void UpdatePropertyData(XmlNode propertyData)
        {
            var countrySettingsNode = propertyData.SelectSingleNode("CountrySettings");

            if (countrySettingsNode != null)
            {
                countrySettingsNode.InnerXml = CountrySettingsInputXml + CountrySettingsOutputXml;
            }
            else
            {
                WriteLog("<CountrySettings> node not found. Searched in: " + propertyData.InnerXml);
            }

            var inputCountrySettingsReferenceNode = propertyData.SelectSingleNode("InputCountrySettingsReference");

            if (inputCountrySettingsReferenceNode != null)
            {
                inputCountrySettingsReferenceNode.InnerText = "United States - Input";
            }
            else
            {
                var xmlNodeTemp = _xmlDocument.CreateNode(XmlNodeType.Element, "InputCountrySettingsReference", null);
                xmlNodeTemp.InnerText = "United States - Input";
                propertyData.AppendChild(xmlNodeTemp);
            }

            var outputCountrySettingsReferenceNode = propertyData.SelectSingleNode("OutputCountrySettingsReference");

            if (outputCountrySettingsReferenceNode != null)
            {
                outputCountrySettingsReferenceNode.InnerText = "United States - Output";
            }
            else
            {
                var xmlNodeTemp = _xmlDocument.CreateNode(XmlNodeType.Element, "OutputCountrySettingsReference", null);
                xmlNodeTemp.InnerText = "United States - Output";
                propertyData.InsertAfter(xmlNodeTemp, inputCountrySettingsReferenceNode);
            }

            var xmlNodesBudgetEntry = propertyData.SelectNodes("BudgetedFinancialData/BudgetEntry");

            if (xmlNodesBudgetEntry != null && xmlNodesBudgetEntry.Count != 0)
            {
                UpdateBudgetedFinancialData(xmlNodesBudgetEntry, propertyData.SelectSingleNode("BudgetedFinancialData"));
            }
        }

        private static void UpdateBaseRentEntry(XmlNodeList baseRentEntries, string subStatus)
        {
            if (subStatus.ToUpper().Equals("CONTRACT"))
            {
                RemoveZeroedAmountBaseEntry(baseRentEntries);
            }

            RemoveDuplicateAmountsBaseEntry(baseRentEntries);
        }

        private static void RemoveDuplicateAmountsBaseEntry(IEnumerable baseRentEntries)
        {
            var amountList = new List<decimal>();

            foreach (XmlNode baseRentEntry in baseRentEntries)
            {
                var xmlNodeRentAmount = baseRentEntry.SelectSingleNode("RentAmount");

                if (xmlNodeRentAmount == null)
                {
                    WriteLog("<RentAmount> node not found. Searched in: " + baseRentEntry.InnerXml);
                    continue;
                }

                decimal tempDecimal;

                if (!decimal.TryParse(xmlNodeRentAmount.InnerText, out tempDecimal)) continue;

                if (!amountList.Contains(tempDecimal))
                {
                    amountList.Add(tempDecimal);
                    continue;
                }

                if (baseRentEntry.ParentNode != null && tempDecimal != 0)
                {
                    baseRentEntry.ParentNode.RemoveChild(baseRentEntry);
                }
            }
        }

        private static void RemoveZeroedAmountBaseEntry(XmlNodeList baseRentEntries)
        {
            var lastNodeBaseRentEntry = baseRentEntries.Item(baseRentEntries.Count - 1);

            if (lastNodeBaseRentEntry == null)
            {
                WriteLog("Error when getting last <BaseRentEntry> node.");
                return;
            }

            var xmlNodeRentAmount = lastNodeBaseRentEntry.SelectSingleNode("RentAmount");

            if (xmlNodeRentAmount == null)
            {
                WriteLog("<RentAmount> node not found. Searched in: " + lastNodeBaseRentEntry.InnerXml);
                return;
            }

            if (xmlNodeRentAmount.InnerText.Equals("0") && lastNodeBaseRentEntry.ParentNode != null)
            {
                lastNodeBaseRentEntry.ParentNode.RemoveChild(lastNodeBaseRentEntry);
            }
        }

        private static void UpdateFreeRentAmount(IEnumerable xmlNodeFreeRents)
        {
            foreach (XmlNode xmlNodeFreeRent in xmlNodeFreeRents)
            {
                var xmlNodeMonthlyAmount = xmlNodeFreeRent.SelectSingleNode("MonthlyAmounts");

                decimal monthlyAmount;

                if (xmlNodeMonthlyAmount == null
                    || string.IsNullOrEmpty(xmlNodeMonthlyAmount.InnerText)
                    || !decimal.TryParse(xmlNodeMonthlyAmount.InnerText, out monthlyAmount))
                {
                    continue;
                }

                xmlNodeMonthlyAmount.InnerText = Math.Round((monthlyAmount / 12), 10).ToString(CultureInfo.InvariantCulture);
            }
        }

        private static void UpdateBudgetedFinancialData(IEnumerable xmlNodesBudgetEntry, XmlNode parentNode)
        {
            var dirInfo = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory + SourceFolder + "\\Templates\\");
            var sourceFileName = ConfigurationManager.AppSettings["AllPossibleAccountsFileName"];
            if (!dirInfo.Exists)
            {
                Console.WriteLine("Source directory (\"" + SourceFolder + "\\Templates\") not found.\n");
                return;
            }
            var fileInfo = new FileInfo(dirInfo.FullName + sourceFileName);
            if (!fileInfo.Exists)
            {
                Console.WriteLine("Source file (\"" + sourceFileName + "\") not found.\n");
                return;
            }
            var xmlDocumentPossibleAccount = new XmlDocument();
            xmlDocumentPossibleAccount.Load(fileInfo.FullName);
            var xmlNodesBaseEntryAll = xmlDocumentPossibleAccount.SelectNodes("reXML/InputAssumptions/PropertyData/BudgetedFinancialData/BudgetEntry");
            if (xmlNodesBaseEntryAll == null || xmlNodesBaseEntryAll.Count == 0)
            {
                Console.WriteLine("Could not found <BudgetEntry>. Attempted path is \"reXML/InputAssumptions/PropertyData/BudgetedFinancialData/BudgetEntry\"\n");
                return;
            }

            var availableEntries = new List<float>();

            var nodesBudgetEntry = xmlNodesBudgetEntry as IList<object> ?? xmlNodesBudgetEntry.Cast<object>().ToList();

            if (xmlNodesBudgetEntry != null)
            {
                availableEntries = nodesBudgetEntry.Cast<XmlNode>().Select(x =>
                {
                    var selectSingleNode = x.SelectSingleNode("BudgetAccountReference");
                    return selectSingleNode != null ? Convert.ToSingle(selectSingleNode.InnerText.ToUpper()) : 0;
                }).ToList();
            }

            var missingBudgetEntries = xmlNodesBaseEntryAll.Cast<XmlNode>()
                .Where(x =>
                {
                    var singleNode = x.SelectSingleNode("BudgetAccountReference");
                    return singleNode != null && !availableEntries.Contains(Convert.ToSingle(singleNode.InnerText));
                })
                .ToList();

            foreach (var missingBudgetEntry in missingBudgetEntries)
            {
                parentNode.AppendChild(_xmlDocument.ImportNode(missingBudgetEntry, true));
            }
        }

        private static void UpdateMakretLeasingAssumptions(XmlNode xmlNodePropertyData)
        {
            var dirInfo = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory + SourceFolder + "\\Templates\\");
            var mlaEntriesFileName = ConfigurationManager.AppSettings["MLAEntriesFileName"];
            if (!dirInfo.Exists)
            {
                Console.WriteLine("Source directory (\"" + SourceFolder + "\\Templates\") not found.\n");
                return;
            }
            var mlaFileInfo = new FileInfo(dirInfo.FullName + mlaEntriesFileName);
            if (!mlaFileInfo.Exists)
            {
                Console.WriteLine("Source file (\"" + mlaEntriesFileName + "\") not found.\n");
                return;
            }

            var xmlDocumentMlaEntries = new XmlDocument();

            xmlDocumentMlaEntries.Load(mlaFileInfo.FullName);

            var xmlNodeMarketLeasingAssumptions = xmlDocumentMlaEntries.SelectSingleNode("/reXML/InputAssumptions/PropertyData/MarketLeasingAssumptions");

            if (xmlNodeMarketLeasingAssumptions == null)
            {
                Console.WriteLine("<MarketLeasingAssumptions> node not found in (\"" + mlaEntriesFileName + "\") file. Attempted path -/reXML/InputAssumptions/PropertyData/MarketLeasingAssumptions.\n");
                return;
            }

            var xmlNodeMarketLeasingAssumptionsTemp = xmlNodePropertyData.SelectSingleNode("MarketLeasingAssumptions");

            if (xmlNodeMarketLeasingAssumptionsTemp == null)
            {
                xmlNodePropertyData.AppendChild(_xmlDocument.ImportNode(xmlNodeMarketLeasingAssumptions, true));
            }
            else
            {
                foreach (XmlNode xmlNode in xmlNodeMarketLeasingAssumptions)
                {

                    xmlNodeMarketLeasingAssumptionsTemp.AppendChild(_xmlDocument.ImportNode(xmlNode, true));
                }
            }
        }

        private static void ApplyMarketLeasingAssumptionsUpdates(string propertyReference)
        {
            if (string.IsNullOrEmpty(propertyReference))
            {
                WriteLog("PropertyReference is empty. Searched in " + _logFileName + ".xml file.");
                return;
            }
            var dirInfo = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory + SourceFolder + "\\Templates\\");
            var mlaUpdatesFileName = ConfigurationManager.AppSettings["MLAUpdatesFileName"];
            if (!dirInfo.Exists)
            {
                Console.WriteLine("Source directory (\"" + SourceFolder + "\\Templates\") not found.\n");
                return;
            }
            var mlaFileInfo = new FileInfo(dirInfo.FullName + mlaUpdatesFileName);
            if (!mlaFileInfo.Exists)
            {
                Console.WriteLine("Source file (\"" + mlaUpdatesFileName + "\") not found.\n");
                return;
            }
            var xmlDocumentMlaUpdates = new XmlDocument();
            xmlDocumentMlaUpdates.Load(mlaFileInfo.FullName);

            var xmlNodeListReXmlEntry = xmlDocumentMlaUpdates.SelectNodes("/reXML/reXML-ENTRY");

            if (xmlNodeListReXmlEntry == null || xmlNodeListReXmlEntry.Count == 0)
            {
                WriteLog("No any <reXML-ENTRY> node found in " + mlaUpdatesFileName + ". Attempted path - /reXML/reXML-ENTRY.");
                return;
            }

            var xmlNodeReXmlEntry = xmlNodeListReXmlEntry.Cast<XmlNode>().Where(x =>
            {
                var selectSingleNode = x.SelectSingleNode("PropertyReference");
                return selectSingleNode != null && selectSingleNode.InnerText.Equals(propertyReference);
            }).FirstOrDefault();

            if (xmlNodeReXmlEntry == null)
            {
                return;
            }

            var xmlNodeNewMarketRentAnnualRate = xmlNodeReXmlEntry.SelectSingleNode("PropertyData/MarketLeasingAssumptions/MarketLeasingAssumptionsTable/MarketRent/NewMarketRentAnnualRate");

            var xmlNodeRenewalMarketRentAnnualRate = xmlNodeReXmlEntry.SelectSingleNode("PropertyData/MarketLeasingAssumptions/MarketLeasingAssumptionsTable/MarketRent/RenewalMarketAnnualRate");

            var xmlNodeListMarketLeaseAssumptionTable = _xmlDocument.SelectNodes("/reXML/InputAssumptions/PropertyData/MarketLeasingAssumptions/MarketLeaseAssumptionTable/MarketRent");

            if (xmlNodeNewMarketRentAnnualRate == null
                || xmlNodeRenewalMarketRentAnnualRate == null
                || xmlNodeListMarketLeaseAssumptionTable == null
                || xmlNodeListMarketLeaseAssumptionTable.Count == 0)
            {
                return;
            }
            try
            {
                foreach (XmlNode xmlNode in xmlNodeListMarketLeaseAssumptionTable)
                {
                    var xmlNode1 = xmlNode.SelectSingleNode("NewMarketRentAnnualRate");
                    if (xmlNode1 != null)
                    {
                        xmlNode1.InnerXml = xmlNodeNewMarketRentAnnualRate.InnerXml;
                    }
                    var xmlNode2 = xmlNode.SelectSingleNode("RenewalMarketRentAnnualRate");

                    if (xmlNode2 != null)
                    {
                        xmlNode2.InnerXml = xmlNodeRenewalMarketRentAnnualRate.InnerXml;
                    }
                }
            }
            catch (Exception ex)
            {
                WriteLog("Error on merging MLAs updates. Error -" + ex);
            }
        }

        private static void WriteLog(string message)
        {
            if (string.IsNullOrEmpty(_logFileName)) return;

            var directoryInfo = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory + "LogFiles\\");
            if (!directoryInfo.Exists) directoryInfo.Create();
            var fileName = directoryInfo.FullName + _logFileName + ".txt";
            try
            {
                using (var sw = new StreamWriter(fileName, true))
                {
                    sw.WriteLine("\n\n");
                    sw.WriteLine(message);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error in writing log. Exception: " + ex.Message);
            }
        }

        private static string GetRandomString(int size = 4)
        {
            var buffer = new char[size];
            for (var i = 0; i < size; i++)
            {
                buffer[i] = Pattern[Random.Next(Pattern.Length)];
            }
            return new string(buffer);
        }

    }
}
