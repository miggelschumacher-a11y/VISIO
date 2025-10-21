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
        Dictionary<string, string> fSpalten = new Dictionary<string, string>();
        List<string> fSpaltenNamen = new List<string>();

        enum TFeldKategorie {
            Normal,
            Pseudo,
            AndereWerte,
            Staerke,
            AnzahlAllePollen,
            AnzahlNektarlose,
            Unbekannt
        }

        private string GetDictValue(string aKey)
        {
            if (fSpalten.ContainsKey(aKey))
                return fSpalten[aKey];
            return "";
        }

        private void SetDictValue(string aKey, string[] aValues, Boolean aIsFloatVal = false)
        {
            int mIndex = fSpaltenNamen.IndexOf(aKey);
            if ((mIndex < 0) || (mIndex >= aValues.Count()))
                return;

            if (fSpalten.ContainsKey(aKey))
            {
                var s = aValues[mIndex];

                if (aIsFloatVal)
                    s = s.Replace('.', ',');
                fSpalten[aKey] = s;
            }
        }
               
        private string[] GetValues(StreamReader aReader)
        {
            // Die Daten sind mit TAB getrennt.
            return aReader.ReadLine().Split('\t');
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

        Boolean FindFeldname(string aFeldName, string aSection)
        {
            Boolean mResult = false;
            List<string> mSection = new List<string>();
            mSection = fIniFile.GetSection(aSection);
            for (int i = 1; i < mSection.Count(); i++)
            {
                string mSectionFeld = mSection[i];
                string[] mSplit = mSectionFeld.Split('=');

                if (mSplit.Count() == 0)
                    return false;

                mSectionFeld = mSplit[0];
                // 05.12.2023
                // Prüfen, ob es reicht, wenn der Feldname aus der Liste nur teilweise in dem zu prüfenden Text (aFeldname) vorhanden ist.
                // Momentan ist das z.B. bei Feldnamen, die mit „Count of“ beginnen so.
                // Deshalb gibt es einen Eintrag „Count of=1“ in der Visio.ini.
                // Die „1“ bestimmt die Teil-Suche.

                // Prüfe, ob aFeldName überhaupt im aktuellen SectionFeld vorkommt.
                if (aFeldName.ToUpper().IndexOf(mSectionFeld.ToUpper(), StringComparison.CurrentCultureIgnoreCase) > -1)
                {
                    // aFeldName kommt im aktuellen Feld vor.
                    // Prüfe, ob der Feldname in mHelpSectionList[i] mit aFeldName vollständig übereinstimmt.
                    // oder, ob es reicht, wenn der Feldname in mHelpSectionList[i] nur zum Teil in aFeldName enthalten ist.
                    if ((aFeldName.ToUpper() == mSectionFeld.ToUpper()) || (fIniFile.GetValue(aSection, mSectionFeld) == "1"))
                        return true;
                }
            }
            return mResult;
        }

        Boolean FeldZulassen(string aFeldName, out TFeldKategorie aKategorie_01)
        {
            aKategorie_01 = TFeldKategorie.Normal;
            if (!FindFeldname(aFeldName, "Relevante-Felder"))
            {
                // Das Feld ist nicht der Liste der relevanten Felder
                if (!FindFeldname(aFeldName, "Nicht-Relevante-Felder"))
                {
                    // Das Feld ist auch nicht in der Liste der unrelevanten Felder
                    // Das bedeutet, das Feld ist unbekannt und wahrscheinlich neu in die Importdatei aufgenommen worden.
                    aKategorie_01 = TFeldKategorie.Unbekannt;
                }
                return false;
            }

            if (FindFeldname(aFeldName, "Struktur"))
                // Wenn es ein Struktur-Feld ist, wird das Feld nicht zugelassen.
                // Die Kategorie kann TPollenKategorie.Normal bleiben, da sie jetzt keine weitere Rolle spielt
                return false;

            if (FindFeldname(aFeldName, "Andere-Werte")) {
                // Wenn es ein AndereWerte-Feld ist, wird das Feld zugelassen und die Kategorie wird auf TPollenKategorie.AndereWerte gesetzt.
                aKategorie_01 = TFeldKategorie.AndereWerte;
                return true;
            }

            if (FindFeldname(aFeldName, "Staerke")) {
                // Wenn es ein Staerke-Feld ist, wird das Feld zugelassen und die Kategorie wird auf TPollenKategorie.Staerke gesetzt.
                aKategorie_01 = TFeldKategorie.Staerke;
                return true;
            }

            if (FindFeldname(aFeldName, "Anzahl-Alle-Pollen")) {
                // Wenn es ein Anzahl-Alle-Pollen-Feld ist, wird das Feld zugelassen und die Kategorie wird auf TPollenKategorie.AnzahlAllePollen gesetzt.
                aKategorie_01 = TFeldKategorie.AnzahlAllePollen;
                return true;
            }

            if (FindFeldname(aFeldName, "Anzahl-Nektarlose")) {
                // Wenn es ein Anzahl-Nektarlose-Feld ist, wird das Feld zugelassen und die Kategorie wird auf TPollenKategorie.AnzahlNektarlose gesetzt.
                aKategorie_01 = TFeldKategorie.AnzahlNektarlose;
                return true;
            }
            return true;
        }

        int IU_VisioImport(DateTime aImportiertWann, string aImportiertAus, string aStudylevel1, string aStudylevel2, string aStudylevel3, string aName, string aImagePath, string aLayerDataPath, Double aStarchPrz, string aPI, int  aID = 0)
        {
            if (aStarchPrz > 999)
                throw new Exception("Starch % darf nicht größer 999 sein!");

            SqlCommand mCmd = new SqlCommand();
            using (SqlConnection mConn = new SqlConnection(fConnectionStr))
            {
                mConn.Open();
                mCmd.Connection = mConn;
                mCmd.CommandType = CommandType.StoredProcedure;
                mCmd.CommandText = "DoVisioImport"; 
                mCmd.Parameters.AddWithValue("@ID", aID);
                mCmd.Parameters.AddWithValue("@ImportiertWann", aImportiertWann);
                mCmd.Parameters.AddWithValue("@ImportiertAus", aImportiertAus);
                mCmd.Parameters.AddWithValue("@Studylevel1", aStudylevel1);
                mCmd.Parameters.AddWithValue("@Studylevel2", aStudylevel2);
                mCmd.Parameters.AddWithValue("@Studylevel3", aStudylevel3);
                mCmd.Parameters.AddWithValue("@Name", aName);
                mCmd.Parameters.AddWithValue("@ImagePath", aImagePath);
                mCmd.Parameters.AddWithValue("@LayerDataPath", aLayerDataPath);
                mCmd.Parameters.AddWithValue("@PI", aPI);
                mCmd.Parameters.AddWithValue("@StarchPrz", aStarchPrz);
                var mReturnParameter = mCmd.Parameters.Add("@ReturnID", SqlDbType.Int);
                mReturnParameter.Direction = ParameterDirection.Output;
                mCmd.ExecuteNonQuery();
                return (int)mReturnParameter.Value;
            }
        }

        int IU_VisioImportPolle(int aVisioImportID,
                                string aVisioName,
                                double aAnzahl,
                                TFeldKategorie aKategorie,
                                int aID = 0)
        {
            SqlCommand mCmd = new SqlCommand();
            using (SqlConnection mConn = new SqlConnection(fConnectionStr))
            {
                mConn.Open();
                mCmd.Connection = mConn;
                mCmd.CommandType = CommandType.StoredProcedure;
                mCmd.CommandText = "IU_VisioImportPolle";
                mCmd.Parameters.AddWithValue("@ID", aID);
                mCmd.Parameters.AddWithValue("@FK_VisioImport", aVisioImportID);
                mCmd.Parameters.AddWithValue("@VisioName", aVisioName);
                mCmd.Parameters.AddWithValue("@Anzahl", aAnzahl);
                mCmd.Parameters.AddWithValue("@Kategorie_01", (int)aKategorie);
                var mReturnParameter = mCmd.Parameters.Add("@ReturnID", SqlDbType.Int);
                mReturnParameter.Direction = ParameterDirection.Output;
                mCmd.ExecuteNonQuery();
                return (int)mReturnParameter.Value;
            }
        }

        void DoImport()
        {
            ReadConfigFiles("SQLSRV", ref fConnectionStr);
            fVisioOrdner = LadeVISIO_ImportOrdner();
            List<string> mDatenZeile = new List<string>();
            fIniFile = new INIFile(@".\VISIO.ini");
            foreach (string fVisioDatenDatei in Directory.EnumerateFiles(fVisioOrdner, "*.tsv", SearchOption.TopDirectoryOnly))
            {
                string mNurDateiName = Path.GetFileName(fVisioDatenDatei);
                string mPiStr = mNurDateiName.Substring(0,10);
                int mID = 0;
                int mVisioImportID = 0;
                using (var mReader = new StreamReader(fVisioDatenDatei)) {
                    var mValues = GetValues(mReader);
                    // In der ersten Zeile sind die Spaltennamen
                    foreach (var mSpalte in mValues)
                    {
                        fSpaltenNamen.Add(mSpalte);
                        fSpalten.Add(mSpalte, "");
                    }

                    // Datenzeile 
                    mValues = GetValues(mReader);
                    SetDictValue("Study level 1", mValues);
                    SetDictValue("Study level 2", mValues);
                    SetDictValue("Study level 3", mValues);
                    SetDictValue("Name", mValues);
                    SetDictValue("Image", mValues);
                    SetDictValue("LayerData", mValues);
                    SetDictValue("Starch %", mValues, true);
                    Double mStarchPrz = 0;
                    Double.TryParse(GetDictValue("Starch %"), out mStarchPrz);

                    mID = IU_VisioImport(DateTime.Now,
                                         mNurDateiName,
                                         GetDictValue("Study level 1"),
                                         GetDictValue("Study level 2"),
                                         GetDictValue("Study level 3"),
                                         GetDictValue("Name"),
                                         GetDictValue("Image"),
                                         GetDictValue("LayerData"),
                                         mStarchPrz,
                                         mPiStr);

                    for (int i = 0; i < mValues.Count(); i++)
                    {
                        //mSpalten[]
                        //mID:= IU_VisioImport(DateTime.Now,
                        //                     Path.GetFileNameWithoutExtension(fVisioDatenDatei),
                        //                     mValues[1].ToString(), // Study level 1

                        //                               fVisioDateiDaten.GetValue('Study level 2', 1),
                        //                               fVisioDateiDaten.GetValue('Study level 3', 1),
                        //                               fVisioDateiDaten.GetValue('Name', 1),
                        //                               fVisioDateiDaten.GetValue('Image', 1),
                        //                               fVisioDateiDaten.GetValue('LayerData', 1),
                        //                               Valr(fVisioDateiDaten.GetValue('starch %', 1)),
                        //                               mPiStr
                        //                               );
                        TFeldKategorie mKategorie_01;
                        int mAnzahl = 0;
                        if ((Int32.TryParse(mValues[i], out mAnzahl)) && (mAnzahl > 0) && (FeldZulassen(fSpaltenNamen[i], out mKategorie_01)))
                        {
                            IU_VisioImportPolle(mVisioImportID,
                                                fSpaltenNamen[i],
                                                mAnzahl,
                                                mKategorie_01);
                        }//if  
                    }//for
                }
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
