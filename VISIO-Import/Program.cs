using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Serialization;
using System.Data;
using System.Data.SqlClient;
using Nocksoft.IO.ConfigFiles;
using MailKit.Net.Smtp;
using MimeKit;
using System.Text;
using System.Timers;

namespace VISIO_Import
{
    #region XMLs
    //
    [Serializable]
    public class DBConnection
    {
        [XmlAttribute("name")]
        public string Name;
        public string DBName;
        public string SourceOrTarget;
        public string ProdServerIP;
        public string ProdSourceString;
        public string DevServerIP;
        public string DevSourceString;

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
        const string cStudy_level_1 = "Study level 1";
        const string cStudy_level_2 = "Study level 2";
        const string cStudy_level_3 = "Study level 3";
        const string cName = "Name";
        const string cImage = "Image";
        const string cLayerData = "LayerData";
        const string cStarchPrz = "Starch %";
        const int cMaxProtkollDateien = 30;
        const int cMaxWarteZeit = 3000;

        INIFile fIniFile;
        string fProtokollDateiname = "";
        string fVisioOrdner;
        string fConnectionStr;
        SqlConnection fConn;
        SqlTransaction fTran;
        readonly Dictionary<string, string> fSpalten = new Dictionary<string, string>();
        readonly List<string> fSpaltenNamen = new List<string>();
        readonly List<string> fProtokoll = new List<string>();
        readonly List<string> fFehler = new List<string>();
        readonly Timer fTimer = new Timer(1000);
        int fWartezeit = cMaxWarteZeit;

        enum TFeldKategorie {
            Normal,
            Pseudo,
            AndereWerte,
            Staerke,
            AnzahlAllePollen,
            AnzahlNektarlose,
            Unbekannt
        }

        public class Analyse
        {
            public int Nummer;
            public string Name;
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
            mCmd.Connection = fConn;
            mCmd.CommandText = @"SELECT VisioOrdner from Konfig";
            SqlDataReader mReader = mCmd.ExecuteReader();
            mReader.Read();
            var s = mReader["VisioOrdner"].ToString();
            mReader.Close();
            return s;
        }

        private Boolean LadeProbePruefauftrag(string aPI, int aUlfd, Analyse aAnalyse)
        {
            SqlCommand mCmd = new SqlCommand();
            mCmd.Connection = fConn;
            mCmd.CommandText = @"select p.Art, d.Oberbegriff
                                 from probe_pruefauftrag p
                                 left join defana d on d.Nummer = p.Art
                                 where p.PiNr = @PI
                                 and p.Ulfd = @Ulfd";

            mCmd.Parameters.AddWithValue("@PI", aPI);
            mCmd.Parameters.AddWithValue("@Ulfd", aUlfd);
            SqlDataReader mReader = mCmd.ExecuteReader();
            mReader.Read();
            try
            {
                if (!mReader.HasRows)
                    return false;
                string s = mReader["Art"].ToString();
                Int32.TryParse(s, out aAnalyse.Nummer);
                aAnalyse.Name = mReader["Oberbegriff"].ToString();
            }
            finally
            {
                mReader.Close();
            }
            return true;
        }

        private Boolean LadePruef(string aPI, int aUlfd, Analyse aAnalyse)
        {
            SqlCommand mCmd = new SqlCommand();
            mCmd.Connection = fConn;
            mCmd.CommandText = @"select p.Analyseart, d.Oberbegriff
                                 from pruef p
                                 left join defana d on d.Nummer = p.Analyseart
                                 where p.PiNr = @PI
                                 and p.Ulfd = @Ulfd";

            mCmd.Parameters.AddWithValue("@PI", aPI);
            mCmd.Parameters.AddWithValue("@Ulfd", aUlfd);
            SqlDataReader mReader = mCmd.ExecuteReader();
            mReader.Read();
            try
            {
                if (!mReader.HasRows)
                    return false;
                string s = mReader["Art"].ToString();
                Int32.TryParse(s, out aAnalyse.Nummer);
                aAnalyse.Name = mReader["Oberbegriff"].ToString();
            }
            finally
            {
                mReader.Close();
            }
            return true;
        }

        private Boolean LadeAnalyseAusProbenAnalyse(string aPI, int aUlfd, Analyse aAnalyse)
        {
            SqlCommand mCmd = new SqlCommand();
            mCmd.Connection = fConn;
            mCmd.CommandText = @"select p.Art, d.Oberbegriff
                                 from ProbenAnalyse_V p
                                 left join defana d on d.Nummer = p.Art
                                 where p.PI = @PI
                                 and p.Ulfd = @Ulfd";

            mCmd.Parameters.AddWithValue("@PI", aPI);
            mCmd.Parameters.AddWithValue("@Ulfd", aUlfd);
            SqlDataReader mReader = mCmd.ExecuteReader();
            mReader.Read();
            try
            {
                if (!mReader.HasRows)
                    return false;
                string s = mReader["Art"].ToString();
                Int32.TryParse(s, out aAnalyse.Nummer);
                aAnalyse.Name = mReader["Oberbegriff"].ToString();
            }
            finally
            {
                mReader.Close();
            }
            return true;
        }

        private void LadeEmailAdressen(List<string> aEmailAdressen)
        {
            using (SqlConnection mConn = new SqlConnection(fConnectionStr))
            {
                mConn.Open();
                SqlCommand mCmd = new SqlCommand();
                mCmd.Connection = mConn;
                mCmd.CommandText = "select e.Email from kundeEmail e where EmailTyp = 'Visio-Import' and aktiv = 1";
                SqlDataReader mReader = mCmd.ExecuteReader();

                while (mReader.Read())
                {
                    aEmailAdressen.Add(mReader["Email"].ToString());
                }
            }
        }

        private Boolean LadeAnalyseAusPerformanceInfo(string aPI, int aUlfd, Analyse aAnalyse)
        {
            SqlCommand mCmd = new SqlCommand();
            mCmd.Connection = fConn;
            mCmd.CommandText = @"select p.AnalyseNr, d.Oberbegriff
                                 from PerformanceInfo_V p
                                 left join defana d on d.Nummer = p.AnalyseNr
                                 where p.PI = @PI
                                 and p.UlfdNr = @Ulfd";

            mCmd.Parameters.AddWithValue("@PI", aPI);
            mCmd.Parameters.AddWithValue("@Ulfd", aUlfd);
            SqlDataReader mReader = mCmd.ExecuteReader();
            mReader.Read();
            try
            {
                if (!mReader.HasRows)
                    return false;
                string s = mReader["Art"].ToString();
                Int32.TryParse(s, out aAnalyse.Nummer);
                aAnalyse.Name = mReader["Oberbegriff"].ToString();
            }
            finally
            {
                mReader.Close();
            }
            return true;
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

        int IU_VisioImport(DateTime aImportiertWann, string aImportiertAus, string aStudylevel1, string aStudylevel2, string aStudylevel3, string aName, string aImagePath, string aLayerDataPath, Double aStarchPrz, string aPI, int aID = 0)
        {
            if (aStarchPrz > 999)
                throw new Exception("Starch % darf nicht größer 999 sein!");

            SqlCommand mCmd = new SqlCommand();
            mCmd.Connection = fConn;
            mCmd.Transaction = fTran;
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

        int IU_VisioImportPolle(int aVisioImportID,
                                string aVisioName,
                                double aAnzahl,
                                TFeldKategorie aKategorie,
                                int aID = 0)
        {
            SqlCommand mCmd = new SqlCommand();
            mCmd.Connection = fConn;
            mCmd.Transaction = fTran;
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

        private void DoProtokoll(List<string> aProtokoll, string aPiStr, string aAnalyseName)
        {
            DateTime mZeitPunkt = DateTime.Now;
            aProtokoll.Add(mZeitPunkt.ToLongDateString() + " " + mZeitPunkt.ToLongTimeString());
            aProtokoll.Add(String.Format("PI: {0} --- Analyse: \"{1}\"", aPiStr, aAnalyseName));
        }

        private void DoSuccess(string aPiStr, string aAnalyseName, string aDateiname)
        {
            List<string> mTemp = new List<string>();
            DoProtokoll(mTemp, aPiStr, aAnalyseName);
            mTemp.Add(String.Format("{0} importiert", aDateiname));
            Console.ForegroundColor = ConsoleColor.Green;
            mTemp.ForEach(x => Console.WriteLine(x));
            fProtokoll.Add("");
            fProtokoll.AddRange(mTemp);
        }

        private void DoError(string aPiStr, string aAnalyseName, string aDateiname, string aFehlerMessage)
        {
            List<string> mTemp = new List<string>();
            DoProtokoll(mTemp, aPiStr, aAnalyseName);
            mTemp.Add("Fehler -> " + aFehlerMessage);
            mTemp.Add(String.Format("{0} nicht importiert", aDateiname));
            mTemp.Add("");
            Console.ForegroundColor = ConsoleColor.Red;
            mTemp.ForEach(x => Console.WriteLine(x));
            fProtokoll.AddRange(mTemp);
            fFehler.AddRange(mTemp);
        }

        private void DoPollenAuftrag(string aPI, int aPollenauftragPiID, int aAnalyseNr, string aBenutzer, int aVisioImportID)
        {
            SqlCommand mCmd = new SqlCommand();
            mCmd.Connection = fConn;
            mCmd.Transaction = fTran;
            mCmd.CommandType = CommandType.StoredProcedure;
            mCmd.CommandText = "DoVisioPolle";
            mCmd.Parameters.AddWithValue("@PI", aPI);
            mCmd.Parameters.AddWithValue("@VisioID", aVisioImportID);
            mCmd.Parameters.AddWithValue("@PollenauftragPiID", aPollenauftragPiID);
            mCmd.Parameters.AddWithValue("@AnalyseNr", aAnalyseNr);
            mCmd.Parameters.AddWithValue("@Benutzer", aBenutzer);
            mCmd.ExecuteNonQuery();
        }

        private int LadePollenAnalysePerPIundAnalyse(string aPI, int aAnalyseNr)
        {
            using (SqlConnection mConn = new SqlConnection(fConnectionStr))
            {
                mConn.Open();
                SqlCommand mCmd = new SqlCommand();
                mCmd.Connection = mConn;
                mCmd.CommandText = @"select ID from PollenanalysePI
                                     where PI = @PI
                                     and AnalyseNr = @AnalyseNr";

                mCmd.Parameters.AddWithValue("@PI", aPI);
                mCmd.Parameters.AddWithValue("@AnalyseNr", aAnalyseNr);
                SqlDataReader mReader = mCmd.ExecuteReader();
                mReader.Read();
                try
                {
                    if (!mReader.HasRows)
                        return 0;
                    Int32.TryParse(mReader["ID"].ToString(), out int mID);
                    return mID;
                }
                finally
                {
                    mReader.Close();
                }
            }
        }

        void DoImport()
        {
            Console.Clear();
            List<string> mImportedFiles = new List<string>();
            DateTime mHeute = DateTime.Today;
            fProtokollDateiname = "Visio-Import-" + mHeute.Day.ToString() + mHeute.Month.ToString() + mHeute.Year.ToString();
            List<string> mDatenZeile = new List<string>();
            fIniFile = new INIFile(@".\VISIO.ini");
            try
            {
                fConnectionStr = ReadConfigFile("SQLSRV");
                using (fConn = new SqlConnection(fConnectionStr))
                {
                    string s = "";
                    fConn.Open();
                    fVisioOrdner = LadeVISIO_ImportOrdner();

                    if (!Directory.Exists(fVisioOrdner))
                    {
                        if (fVisioOrdner.Trim() == "")
                            s = "Visio-Ordner ist nicht gesetzt!";
                        else
                            s = String.Format("Visio-Ordner ({0}) für Importdaten nicht gefunden!", fVisioOrdner);

                        fProtokoll.Add(s);
                        fFehler.Add(s);
                        return;
                    }

                    foreach (string fVisioDatenDatei in Directory.EnumerateFiles(fVisioOrdner, "*.tsv", SearchOption.TopDirectoryOnly))
                    {
                        string mNurDateiName = Path.GetFileName(fVisioDatenDatei);
                        string mPiStr = mNurDateiName.Substring(0, 10);
                        s = mNurDateiName.Substring(10, 2);
                        Int32.TryParse(s, out int mUlfd);
                        Analyse mAnalyse = new Analyse();
                        // Versuche die Analyse-Daten verschiedenen Tabellen zu laden
                        if (!(   // Zunächst in der Tabelle ProbenAnalysen suchen
                                 LadeAnalyseAusProbenAnalyse(mPiStr, mUlfd, mAnalyse)
                              // Wenn nichts gefunden wurde, in der Tabelle PerformanceInfo suchen
                              || LadeAnalyseAusPerformanceInfo(mPiStr, mUlfd, mAnalyse)
                              // Wenn nichts gefunden wurde, in der Tabelle Proben_Pruefauftrag suchen
                              || LadeProbePruefauftrag(mPiStr, mUlfd, mAnalyse)
                              // Wenn nichts gefunden wurde, in der Tabelle Pruef suchen
                              || LadePruef(mPiStr, mUlfd, mAnalyse)
                           ))
                        {
                            // In keine der oben genannten Tabelle konnte ein Datensatz mit "mPiStr" und "mUlfd" gefunden werden 
                            DoError(mPiStr,
                                     "?",// Analysename
                                     mNurDateiName,
                                     String.Format("Konnte Probenanalyse mit PI{0} und Lfd-Nr. {1} nicht finden!", mPiStr, mUlfd));
                            continue;
                        }

                        int mVisioImportID = 0;
                        using (var mReader = new StreamReader(fVisioDatenDatei))
                        {
                            fSpalten.Clear();
                            var mValues = GetValues(mReader);
                            // In der ersten Zeile sind die Spaltennamen
                            foreach (var mSpalte in mValues)
                            {
                                fSpaltenNamen.Add(mSpalte);
                                fSpalten.Add(mSpalte, "");
                            }

                            // In der zweiten Zeile sind die Daten
                            mValues = GetValues(mReader);
                            SetDictValue(cStudy_level_1, mValues);
                            SetDictValue(cStudy_level_2, mValues);
                            SetDictValue(cStudy_level_3, mValues);
                            SetDictValue(cName, mValues);
                            SetDictValue(cImage, mValues);
                            SetDictValue(cLayerData, mValues);
                            SetDictValue(cStarchPrz, mValues, true);
                            Double mStarchPrz = 0;
                            Double.TryParse(GetDictValue(cStarchPrz), out mStarchPrz);

                            fTran = fConn.BeginTransaction();
                            try
                            {
                                List<string> mUnbekannteFelder = new List<string>();
                                mVisioImportID = IU_VisioImport(DateTime.Now,
                                                                mNurDateiName,
                                                                GetDictValue(cStudy_level_1),
                                                                GetDictValue(cStudy_level_2),
                                                                GetDictValue(cStudy_level_3),
                                                                GetDictValue(cName),
                                                                GetDictValue(cImage),
                                                                GetDictValue(cLayerData),
                                                                mStarchPrz,
                                                                mPiStr);

                                for (int i = 0; i < mValues.Count(); i++)
                                {
                                    TFeldKategorie mKategorie_01 = TFeldKategorie.Normal;
                                    int mAnzahl = 0;
                                    if (Int32.TryParse(mValues[i], out mAnzahl) && (mAnzahl > 0))
                                    {
                                        if (FeldZulassen(fSpaltenNamen[i], out mKategorie_01))
                                            IU_VisioImportPolle(mVisioImportID,
                                                                fSpaltenNamen[i],
                                                                mAnzahl,
                                                                mKategorie_01);
                                        else if (mKategorie_01 == TFeldKategorie.Unbekannt)
                                            mUnbekannteFelder.Add(String.Format("Unbekanntes Feld \"{0}\"", fSpaltenNamen[i]));
                                    }
                                }//for

                                if (mUnbekannteFelder.Count > 0)
                                    throw new Exception(mUnbekannteFelder.ToString());

                                // MasterID für den Pollenauftrag laden
                                int mPollenauftragPiID = LadePollenAnalysePerPIundAnalyse(mPiStr, mAnalyse.Nummer);
                                // Pollenauftrag erzeugen
                                DoPollenAuftrag(mPiStr,
                                                mPollenauftragPiID,
                                                mAnalyse.Nummer,
                                                "Server",
                                                mVisioImportID);


                                fTran.Commit();
                                // Die importierten Dateien sollen gelöscht werden, aber an dieser Stelle ruft das Löschen eine Exception hervor!
                                // Daher werden sie Dateien hier gesammelt, um sie später zu löschen.
                                mImportedFiles.Add(fVisioDatenDatei);
                                DoSuccess(mPiStr,
                                          mAnalyse.Name,
                                          mNurDateiName);
                            }
                            catch (Exception Ex)
                            {
                                fTran.Rollback();
                                DoError(mPiStr,
                                        mAnalyse.Name,
                                        mNurDateiName,
                                        Ex.Message);
                            }
                        }
                    }
                }
                // Importierte Dateien löschen
                foreach (var mFile in mImportedFiles)
                {
                    try
                    {
                        File.Delete(mFile);
                    }
                    catch (Exception Ex)
                    {
                        var s = String.Format("Fehler {0} --- Datei: \"{1}\" konnte nach dem Import nicht gelöscht werden!", mFile, Ex.Message);
                        fProtokoll.Add(s);
                        fFehler.Add(s);
                    }
                }
            }
            finally
            {
                Console.ForegroundColor = ConsoleColor.Gray;
                if (fProtokoll.Count() > 0)
                {
                    string mProtokollDir = fVisioOrdner + "\\Protokoll\\";
                    string s = mProtokollDir + fProtokollDateiname + ".prot";

                    if (File.Exists(s))
                        File.AppendAllLines(s, fProtokoll);
                    else
                        File.WriteAllLines(s, fProtokoll);

                    var mCreationDateList = Directory.EnumerateFiles(mProtokollDir, "*.prot", SearchOption.TopDirectoryOnly).ToDictionary(x => File.GetCreationTime(x), x => x);

                    if (mCreationDateList.Count > cMaxProtkollDateien)
                    {
                        foreach (KeyValuePair<DateTime, string> mFile in mCreationDateList.OrderBy(x => x.Key))
                        {
                            File.Delete(mFile.Value);
                            mCreationDateList.Remove(mFile.Key);

                            if (mCreationDateList.Count <= cMaxProtkollDateien)
                                break;
                        }
                    }
                }

                if (fFehler.Count > 0)
                {
                    List<string> mEmailAdressen = new List<string>();
                    var message = new MimeMessage();
                    message.From.Add(new MailboxAddress("VISIO-Import-App", "Ladis@intertek.com"));

                    var mTestModus = ReadConfigFile("Test-Modus");
                    if (mTestModus.Trim() == "1")
                        // Programm läuft laut INI-File im Test-Modus
                        message.To.Add(new MailboxAddress("VISIO-Fehler-Empfänger", "michael.schumacher@intertek.com"));
                    else
                    {
                        // Programm läuft laut INI-File NICHT im Test-Modus
                        LadeEmailAdressen(mEmailAdressen);
                        // Für jeden Eintrag in mEmailAdressen eine MailboxAddress erzeugen und an message.To anhängen
                        mEmailAdressen.ForEach(x => message.To.Add(new MailboxAddress("VISIO-Fehler-Empfänger", x)));
                    }
                    
                    message.Subject = "VISIO-Import-Fehler";

                    // Bodytext zusammenbauen
                    var mBuilder = new BodyBuilder();
                    StringBuilder mText = new StringBuilder();
                    // Alle Fehlerzeilen in mText kopieren
                    fFehler.ForEach(x => mText.AppendLine(x));
                    var st = mText.ToString();
                    // mText in mBuilder.TextBody kopieren
                    mBuilder.TextBody = mText.ToString();
                    // Texte aus mBuilder message.Body kopieren
                    message.Body = mBuilder.ToMessageBody();

                    using (var client = new SmtpClient())
                    {
                        var mHost = ReadConfigFile("Host");
                        if (mHost == "")
                            mHost = "10.135.26.8";

                        var mPort = ReadConfigFile("Port");
                        Int32.TryParse(mPort, out int mIntPort);

                        if (mIntPort  == 0)
                            mIntPort = 25;

                        Console.WriteLine("Versende E-Mails");
                        client.Connect(mHost, mIntPort, false);
                        client.Send(message);
                        client.Disconnect(true);
                    }
                }
            }
            Console.Write("Programm wird beendet");
            WriteWarteInfo();
            fTimer.Elapsed += OnTimedEvent;
            fTimer.Enabled = true;
            fWartezeit -= 1000;
            do {} while (fTimer.Enabled);
        }

        private void WriteWarteInfo()
        {
            Console.Write(".");
        }

        private void OnTimedEvent(Object source, ElapsedEventArgs e)
        { 
            WriteWarteInfo();
            fWartezeit -= 1000;
            fTimer.Enabled = (fWartezeit >= 0);
        }

        #region Config Files
        private string ReadConfigFile(string aSuchPara)
        {
            var mSerializer = new XmlSerializer(typeof(ConfigContainer));
            using (var mStream = new FileStream(@".\VISIO_Config.xml", FileMode.Open))
            {
                var mContainer = mSerializer.Deserialize(mStream) as ConfigContainer;
                mStream.Close();

                List<DBConnection> dbcSrc = mContainer.DBConnections.FindAll(x => x.Name.Equals(aSuchPara));
                if (dbcSrc.Count > 0)
                {
                    AppParameter mAppParameter = mContainer.AppParameters.Find(x => x.Name.Equals("Test-Modus"));
                    Boolean mTestModus = (mAppParameter.Value == "1");
                    if (mTestModus)
                        return dbcSrc[0].DevSourceString.Trim();
                    else
                        return dbcSrc[0].ProdSourceString.Trim();
                }
                else
                {
                    AppParameter mAppParameter = mContainer.AppParameters.Find(x => x.Name.Equals(aSuchPara));
                    if (mAppParameter != null)
                        return mAppParameter.Value;

                }
                return "";
            }
        }
        #endregion
        static void Main(string[] args)
        {
            new Program().DoImport();
        }

    }
}
