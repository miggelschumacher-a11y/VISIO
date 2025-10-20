using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;
using System.Data;
using System.Data.SqlClient;
using Nocksoft.IO.ConfigFiles;

namespace VISIO_Import
{
        #region XMLs
        //
        [Serializable]
        public class DBConnection
        {
            [XmlAttribute("name")]
            public string Name;

            public string SourceOrTarget;

            public string ServerIP;

            public string DBName;

            public string ProdSourceString;

            public string DevSourceString;

            public string Modus;
        }
        public class AppParameter
        {
            [XmlAttribute("Name")]
            public string Name;

            [XmlAttribute("Value")]
            public string Value;
        }

        [XmlRoot("ConfigCollection")]
        public class ConfigContainer
        {
            [XmlArray("DBConnections"), XmlArrayItem("DBConnection")]
            public List<DBConnection> DBConnections = new List<DBConnection>();

            [XmlArray("AppParameters"), XmlArrayItem("AppParameter")]
            public List<AppParameter> AppParameters = new List<AppParameter>();
        }
    #endregion
    class Program
    {
        INIFile fIniFile;
        string fVisioOrdner;
        string fConnectionStr;

        enum TFeldKategorie {
            Normal,
            Pseudo,
            AndereWerte,
            Staerke,
            AnzahlAllePollen,
            AnzahlNektarlose,
            Unbekannt
        }

        private string LadeVISIO_ImportOrdner()
        {
            SqlCommand mCmd = new SqlCommand();
            using (SqlConnection mConn = new SqlConnection(fConnectionStr))
            {
                mConn.Open();
                mCmd.Connection = mConn;
                mCmd.CommandText = "SELECT VisioOrdner from Konfig";
                SqlDataReader mReader = mCmd.ExecuteReader();
                mReader.Read();
                return mReader["VisioOrdner"].ToString();
            }
        }

        Boolean FindFeldname(string aFeldName, string aSection){
            return true;
        }

        Boolean FeldZulassen(string aFeldName, out TFeldKategorie aKategorie_01)
        {
            aKategorie_01 = TFeldKategorie.Unbekannt;
            return true;
        }

        int IU_VisioImportPolle(int aVisioImportID,
                                string aVisioName,
                                double aAnzahl,
                                int aKategorie,
                                int aID = 0)
        {
            return 0;
        }

        void DoImport()
        {
            ReadConfigFiles("SQLSRV", ref fConnectionStr);
            fVisioOrdner = LadeVISIO_ImportOrdner();
            List<string> mDatenZeile = new List<string>();
            fIniFile = new INIFile(@".\VISIO.ini");
            foreach (string fVisioDatenDatei in Directory.EnumerateFiles(fVisioOrdner, "*.tsv", SearchOption.TopDirectoryOnly))
            {
                int mZeileNr = -1;
                int mVisioImportID = 0;
                
                List<string> mSpaltenNamen = new List<string>();
                using (var mReader = new StreamReader(fVisioDatenDatei)) {
                    while (!mReader.EndOfStream)
                    {
                        mZeileNr++;
                        // In der ersten Zeile sind die Spaltennamen
                        // Die Daten sind mit TAB getrennt.
                        var mValues = mReader.ReadLine().Split('\t');

                        if (mZeileNr == 0)
                        {
                            foreach (var mSpalte in mValues)
                                mSpaltenNamen.Add(mSpalte);
                            continue;
                        }

                        for (int i = 0; i < mValues.Count(); i++)
                        {
                            TFeldKategorie mKategorie_01;
                            int mAnzahl = 0;
                            if ((Int32.TryParse(mValues[i], out mAnzahl)) && (mAnzahl > 0) && (FeldZulassen(mSpaltenNamen[i], out mKategorie_01)))
                            {
                                IU_VisioImportPolle(mVisioImportID,
                                                    mSpaltenNamen[i],
                                                    mAnzahl,
                                                    (int)mKategorie_01
                                                   );
                            }//if  
                        }//for
                    }
                    
                }
                    //string s = fVisioDatenDatei.R.GetValue(fVisioDatenDatei.Fieldnames[j], 1));
                //string s = Trim(fVisioDateiDaten.GetValue(fVisioDateiDaten.Fieldnames[j], 1));
                //mAnzahl:= 0;
                //TryStrToFloat(s, mAnzahl);

            }
        }

        #region Config Files
        private void ReadConfigFiles(string aSuch, ref string aResult)
        {
            try
            {
                aResult = "";
                var mSerializer = new XmlSerializer(typeof(ConfigContainer));
                using (var mStream = new FileStream(@".\VISIO_Config.xml", FileMode.Open))
                {
                    var mContainer = mSerializer.Deserialize(mStream) as ConfigContainer;
                    mStream.Close();
                    //
                    List<DBConnection> dbcSrc = mContainer.DBConnections.FindAll(x => x.Name.Equals(aSuch));
                    //
                    if (dbcSrc.Count > 0)
                    {
                        if (dbcSrc[0].Modus == "Entwicklung")
                            aResult = dbcSrc[0].DevSourceString.Trim();
                        else
                            aResult = dbcSrc[0].ProdSourceString.Trim();
                    }
                    else
                    {
                        AppParameter mAppParameter = mContainer.AppParameters.Find(x => x.Name.Equals(aSuch));
                        if (mAppParameter != null)
                            aResult = mAppParameter.Value;

                    }
                }
            }
            catch (Exception ex)
            {
                aResult = @"Data Source=EDEUBREAPP003\SQLBRE03;Initial Catalog=ladisInSQL;User ID=Ladis;Password=Winter2015!";
            }
        }
        private string ReadConfigValue(string pName)
        {
            string retval = "";
            try
            {
                var serializer = new XmlSerializer(typeof(ConfigContainer));

                var stream = new FileStream(@".\VISIO_Config.xml", FileMode.Open);
                var container = serializer.Deserialize(stream) as ConfigContainer;
                stream.Close();
                //
                List<AppParameter> appParam = container.AppParameters.FindAll(x => x.Name.Equals(pName));
                //
                if (appParam.Count > 0)
                {
                    retval = appParam[0].Value.Trim();
                }
            }
            catch (Exception ex)
            {
                retval = "";
            }
            return retval;
        }
        #endregion
        static void Main(string[] args)
        {
            new Program().DoImport();
        }

    }
}
