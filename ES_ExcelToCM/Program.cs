using System;
using System.Data;
using System.IO;
using System.Xml;
using HP.HPTRIM.SDK;
using OfficeOpenXml;

namespace ES_ExcelToCM
{
    internal class Program
    {
        static public Database db;
        static String filepath = "", dataset = "", port = "", ipaddress = "";

        static void Main(string[] args)
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load("config.xml");
                XmlNode configNode = xmlDoc.SelectSingleNode("//config");

                filepath = configNode.SelectSingleNode("filepath").Attributes["path"].Value;
                dataset = configNode.SelectSingleNode("dataset").Attributes["ID"].Value;
                port = configNode.SelectSingleNode("port").Attributes["value"].Value;
                ipaddress = configNode.SelectSingleNode("ipaddress").Attributes["value"].Value;
                ConnectDb();

                var dt = new DataTable();
                using (var package = new ExcelPackage(new FileInfo(filepath)))
                {
                    var worksheet = package.Workbook.Worksheets[0];
                    for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                    {
                        dt.Columns.Add(worksheet.Cells[1, col].Value.ToString());
                    }
                    for (int row = 2; row <= worksheet.Dimension.Rows; row++)
                    {
                        var dataRow = dt.NewRow();
                        for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                        {
                            dataRow[col - 1] = worksheet.Cells[row, col].Value;
                        }
                        dt.Rows.Add(dataRow);
                    }
                }
                for (int i = 2; i < dt.Rows.Count; i++)
                {
                    var search = new TrimMainObjectSearch(db, BaseObjectTypes.Location);
                    search.SetSearchString("id:" + dt.Rows[i]["GFIS ID"] + "-" + dt.Rows[i]["Eng / Projet ID"]);
                    if (search.Count == 0)
                    {
                        int flag = 0;
                        if (dt.Rows[i]["Legal Entity_descr"] + "" == "E&Y Corporate Finance" || dt.Rows[i]["Legal Entity_descr"].ToString() == "E&Y Societe Avocats" || dt.Rows[i]["Legal Entity_descr"].ToString() == "EY Tax and Law France" || dt.Rows[i]["Legal Entity_descr"].ToString() == "Fabernovel" || dt.Rows[i]["Legal Entity_descr"].ToString() == "Fabernovel Group" || dt.Rows[i]["Legal Entity_descr"].ToString() == "VENTURY AVOCATS")
                        {
                            flag = 1;
                        }
                        if (dt.Rows[i]["BU_Descr"].ToString() == "E&Y Corporate Finance" || dt.Rows[i]["BU_Descr"].ToString() == "E&Y Societe Avocats" || dt.Rows[i]["BU_Descr"].ToString() == "EY Societe Avocats" || dt.Rows[i]["BU_Descr"].ToString() == "Fabernovel" || dt.Rows[i]["BU_Descr"].ToString() == "Fabernovel Group" || dt.Rows[i]["BU_Descr"].ToString() == "MAB - E&Y Societe Avocats" || dt.Rows[i]["OU_Descr"].ToString() == "MAB - VENTURY Avocats" || dt.Rows[i]["OU_Descr"].ToString() == "VENTURY Avocats")
                        {
                            flag = 1;
                        }
                        if (dt.Rows[i]["OU_Descr"].ToString() == "BTS - FraLux" || dt.Rows[i]["OU_Descr"].ToString() == "Business Law - FraLux" || dt.Rows[i]["OU_Descr"].ToString() == "Corporate TAX Paris - FR" || dt.Rows[i]["OU_Descr"].ToString() == "Financial Services" || dt.Rows[i]["OU_Descr"].ToString() == "FISCAL" || dt.Rows[i]["OU_Descr"].ToString() == "FSO TAS - FR" || dt.Rows[i]["OU_Descr"].ToString() == "FSO TAX - FR" || dt.Rows[i]["OU_Descr"].ToString() == "HCP - FraLux" || dt.Rows[i]["OU_Descr"].ToString() == "Ind. Tax - FraLux" || dt.Rows[i]["OU_Descr"].ToString() == "IT" || dt.Rows[i]["OU_Descr"].ToString() == "ITS - FraLux" || dt.Rows[i]["OU_Descr"].ToString() == "ITS Desk - FR" || dt.Rows[i]["OU_Descr"].ToString() == "JF" || dt.Rows[i]["OU_Descr"].ToString() == "Law Employment - FR" || dt.Rows[i]["OU_Descr"].ToString() == "Legal" || dt.Rows[i]["OU_Descr"].ToString() == "PAS ETC" || dt.Rows[i]["OU_Descr"].ToString() == "PAS FSO-France" || dt.Rows[i]["OU_Descr"].ToString() == "PAS-France" || dt.Rows[i]["OU_Descr"].ToString() == "Tax Management" || dt.Rows[i]["OU_Descr"].ToString() == "Tax Regions - FR" || dt.Rows[i]["OU_Descr"].ToString() == "Transaction Integration - FR" || dt.Rows[i]["OU_Descr"].ToString() == "Transaction Tax - FraLux" || dt.Rows[i]["OU_Descr"].ToString() == "Transverse JF" || dt.Rows[i]["OU_Descr"].ToString() == "TS - FR")
                        {
                            flag = 1;
                        }

                        if (flag == 1)
                        {
                            Location org = new Location(db);
                            org.TypeOfLocation = LocationType.Organization;

                            FieldDefinition fd = new FieldDefinition(db, "Client Code");
                            UserFieldValue udf = new UserFieldValue(dt.Rows[i]["GFIS ID"] + "");
                            org.SetFieldValue(fd, udf);

                            fd = new FieldDefinition(db, "Client ID GFIS");
                            udf = new UserFieldValue(dt.Rows[i]["GFIS ID"] + "");
                            org.SetFieldValue(fd, udf);

                            fd = new FieldDefinition(db, "Client Name");
                            udf = new UserFieldValue(dt.Rows[i]["GFIS Name"] + "");
                            org.SetFieldValue(fd, udf);

                            fd = new FieldDefinition(db, "Client Status");
                            udf = new UserFieldValue(dt.Rows[i]["Eng / Projet ID satus descr"] + "");
                            org.SetFieldValue(fd, udf);

                            fd = new FieldDefinition(db, "Company Code Name");
                            udf = new UserFieldValue(dt.Rows[i]["Legal Entity_descr"] + "");
                            org.SetFieldValue(fd, udf);

                            fd = new FieldDefinition(db, "Engagement Code");
                            udf = new UserFieldValue(dt.Rows[i]["Eng / Projet ID"] + "");
                            org.SetFieldValue(fd, udf);

                            fd = new FieldDefinition(db, "Engagement ID GFIS");
                            udf = new UserFieldValue(dt.Rows[i]["Eng / Projet ID"] + "");
                            org.SetFieldValue(fd, udf);

                            fd = new FieldDefinition(db, "Engagement Name");
                            udf = new UserFieldValue(dt.Rows[i]["Eng / Projet Name"] + "");
                            org.SetFieldValue(fd, udf);

                            fd = new FieldDefinition(db, "Engagement Status");
                            udf = new UserFieldValue(dt.Rows[i]["Eng / Projet ID satus descr"] + "");
                            org.SetFieldValue(fd, udf);

                            fd = new FieldDefinition(db, "Manager GPN");
                            udf = new UserFieldValue(dt.Rows[i]["Eng Manager"] + "");
                            org.SetFieldValue(fd, udf);

                            fd = new FieldDefinition(db, "Manager Name");
                            udf = new UserFieldValue(dt.Rows[i]["Eng Manager NM"] + "");
                            org.SetFieldValue(fd, udf);

                            fd = new FieldDefinition(db, "Partner GPN");
                            udf = new UserFieldValue(dt.Rows[i]["Eng Partner"] + "");
                            org.SetFieldValue(fd, udf);

                            fd = new FieldDefinition(db, "Partner Name");
                            udf = new UserFieldValue(dt.Rows[i]["Eng Partner NM"] + "");
                            org.SetFieldValue(fd, udf);

                            fd = new FieldDefinition(db, "Business Unit");
                            udf = new UserFieldValue(dt.Rows[i]["BU_Descr"] + "");
                            org.SetFieldValue(fd, udf);

                            org.IdNumber = dt.Rows[i]["GFIS ID"] + "-" + dt.Rows[i]["Eng / Projet ID"];
                            org.SortName = dt.Rows[i]["GFIS Name"] + "-" + dt.Rows[i]["Eng / Projet Name"];

                            org.Save();
                            Console.WriteLine(org.FormattedName + " created.");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            DisconnectDb();
            Console.ReadLine();
        }

        static public void ConnectDb()
        {
            db = new Database();
            db.Id = dataset;
            db.WorkgroupServerPort = Convert.ToInt32(port);
            db.WorkgroupServerName = ipaddress;
            db.Connect();
            Console.WriteLine("Connected to Database");
        }

        static public void DisconnectDb()
        {
            if (db != null)
            {
                db.Disconnect();
            }
        }
    }
}