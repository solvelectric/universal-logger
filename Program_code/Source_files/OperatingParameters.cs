using OfficeOpenXml; // Az xlsx fájlok kezeléséhez szükséges
using System;
using System.Collections.Generic;
using System.IO; // A fájlkezeléshez szükséges
using System.Linq;
using System.Net.Mail; // Az e-mail címek kezelése miatt kell
using System.Text;
using System.Text.RegularExpressions; // Az e-mail címek helyességének vizsgálatához kell

namespace universal_logger
{
    /// <summary>A paramétertáblázatot kezelő osztály</summary>
    /// <remarks>
    /// Ezzel az osztállyal lehet beolvastatni a paramétertáblázatot, valamint lekérdezni az egyes működési paramétereket.
    /// Az osztály nem rendelkezik kivételkezeléssel, így ezt felsőbb szinten kell megoldani.
    /// </remarks>
    class OperatingParameters
    {
        /// <summary>Ez az osztály tartalmazza a táblázatból lekérdezhető programbeállításokat</summary>
        public class SettingsData
        {
            /// <summary>Konstruktor</summary>
            public SettingsData()
            {
                emailList = new MailAddressCollection(); // Inicializáljuk az e-mail címeket tartalmazó objektumot
                useDeviceDateTime = false; // Alapértelmezésként a pc idejét használjuk
            }

            /// <summary>Az e-mail címeket tartalmazó objektum</summary>
            public MailAddressCollection emailList;

            /// <summary>A használandó COM port neve</summary>
            public string port;

            /// <summary>A használandó baud rate</summary>
            public int baudRate;

            /// <summary>A kommunikáció során használt időtúllépés ms-ban</summary>
            public int timeout;

            /// <summary>A rendszer azonosítószáma</summary>
            public ushort systemId;

            /// <summary>A rendszer megnevezése</summary>
            public string systemName;

            /// <summary>Ebben az órában küldjük el az e-mail címekre a napi fájlokat tartalmazó jelentést</summary>
            public ushort emailSendingHour;

            /// <summary>A logfájlok könyvtárának a neve</summary>
            public string logDirectoryName;

            /// <summary>A rendszerben található lekérdező parancsok száma</summary>
            /// <remarks>Nem tartalmazza az általános parancs protokollban meghatározott kötelező parancsokat.</remarks>
            public byte numOfQueryCommands;

            /// <summary>A rendszerben található lekérdezhető adatok száma</summary>
            /// <remarks>Nem tartalmazza az általános parancs protokollban meghatározott kötelező adatokat.</remarks>
            public ushort numOfQuerySections;

            /// <summary>A rendszerben található lekérdezhető bitek száma</summary>
            /// <remarks>Nem tartalmazza az általános parancs protokollban meghatározott kötelező biteket.</remarks>
            public int numOfQueryBits;

            /// <summary>A rendszerben található beállító parancsok száma</summary>
            /// <remarks>Nem tartalmazza az általános parancs protokollban meghatározott kötelező parancsokat.</remarks>
            public byte numOfSetCommands;

            /// <summary>A rendszerben található beállítható adatok száma</summary>
            /// <remarks>Nem tartalmazza az általános parancs protokollban meghatározott kötelező adatokat.</remarks>
            public ushort numOfSetSections;

            /// <summary>A rendszerben található beállítható bitek száma</summary>
            /// <remarks>Nem tartalmazza az általános parancs protokollban meghatározott kötelező biteket.</remarks>
            public int numOfSetBits;

            /// <summary>Az időzítő automatikus indítását jelző változó</summary>
            /// <remarks>
            /// A lehetséges értékek, és azok jelentése:
            /// 0  - nem kell automatikusan elindítani az időzítőt
            /// 1  - az időzítőt automatikusan el indítani, az intervallum 1s
            /// 2  - az időzítőt automatikusan el indítani, az intervallum 2s
            /// 3  - az időzítőt automatikusan el indítani, az intervallum 5s
            /// 4  - az időzítőt automatikusan el indítani, az intervallum 10s
            /// 5  - az időzítőt automatikusan el indítani, az intervallum 20s
            /// 6  - az időzítőt automatikusan el indítani, az intervallum 30s
            /// 7  - az időzítőt automatikusan el indítani, az intervallum 1min
            /// 8  - az időzítőt automatikusan el indítani, az intervallum 2min
            /// 9  - az időzítőt automatikusan el indítani, az intervallum 5min
            /// 10 - az időzítőt automatikusan el indítani, az intervallum 10min
            /// 11 - az időzítőt automatikusan el indítani, az intervallum 30min
            /// 12 - az időzítőt automatikusan el indítani, az intervallum 60min
            /// </remarks>
            public byte timerAutoStart;

            /// <summary>A rendszerben lévő lekérdező parancsok listája</summary>
            /// <remarks>Nem tartalmazza az általános parancs protokollban meghatározott kötelező parancsokat.</remarks>
            public List<QueryCommand> queryCommands;

            /// <summary>A rendszerben lévő beállító parancsok listája</summary>
            /// <remarks>Nem tartalmazza az általános parancs protokollban meghatározott kötelező parancsokat.</remarks>
            public List<SetCommand> setCommands;

            /// <summary>Igaz, ha az eszköz küldi el a mentéshez használt dátumot és időt</summary>
            public bool useDeviceDateTime;

            /// <summary>A logfájl fejléce</summary>
            public string logFileHeader;
        }

        /// <summary>Ez az osztály tartalmazza a lekérdező parancsok leírását</summary>
        public class QueryCommand
        {
            /// <summary>Konstruktor</summary>
            /// <param name="numberOfSections">A parancsra kapott válasz ennyi adatot tartalmaz</param>
            public QueryCommand(byte numberOfSections)
            {
                numOfSections = numberOfSections; // Lementjük az adatok számát,
                sections = new List<QuerySection>(numberOfSections); // és inicialiáljuk a lista méretét is
                return;
            }

            /// <summary>A parancs kódja</summary>
            public byte code;

            /// <summary>A parancs neve</summary>
            public string name;

            /// <summary>A parancsra kapott válasz adatbájtjainak (DATAn) a száma</summary>
            public byte dataBytes;

            /// <summary>A parancsra kapott válasz ennyi adatot tartalmaz</summary>
            public byte numOfSections;

            /// <summary>Igaz, ha a parancs melletti, időzítővel való lekérdezéshez tartozó jelölőnégyzetet be kell jelölni</summary>
            public bool autoQuery;

            /// <summary>A parancsra kapott válasz adatainak listája</summary>
            public List<QuerySection> sections;
        }

        /// <summary>Ez az osztály tartalmazza a lekérdező parancsok adatainak leírását</summary>
        public class QuerySection
        {
            /// <summary>Konstruktor</summary>
            /// <param name="numberOfBits">Hexadecimális érték esetén a bitek száma</param>
            public QuerySection(byte numberOfBits)
            {
                numOfBits = numberOfBits; // Lementjük a bitek számát,
                bits = new List<QueryBit>(numberOfBits); // és inicialiáljuk a lista méretét is
                return;
            }

            /// <summary>Az adat neve</summary>
            public string name;

            /// <summary>Amennyiben az adat hexadecimális szám, akkor ez az érték mondja meg, hogy hány bitet tartalmaz</summary>
            public byte numOfBits;

            /// <summary>Az adat típusa</summary>
            public Type type;

            /// <summary>Az adat tizedesjegyeinek a száma</summary>
            /// <remarks>
            /// Ha az érték 0, akkor az adat egész szám.
            /// Ha az érték -1, akkor az adat dátum/idő
            /// Ha az érték -2, akkor az adat hexadecimális szám.
            /// </remarks>
            public sbyte decimals;

            /// <summary>Ha igaz, akkor az adatot le kell menteni a log fájlba</summary>
            public bool save;

            /// <summary>Amennyiben az adat hexadecimális szám, akkor ez a lista tartalmazza a biteket</summary>
            public List<QueryBit> bits;

            /// <summary>Az adat típusának a beállítása</summary>
            /// <param name="dataType">Az adat típusa</param>
            public void SetType(Object dataType)
            {
                type = dataType.GetType(); // Lementjük az adat típusát
                return;
            }
        }

        /// <summary>Ez az osztály tartalmazza a lekérdező parancsok bitjeinek leírását</summary>
        public class QueryBit
        {
            /// <summary>A bit neve</summary>
            public string name;

            /// <summary>Igaz, ha a bit használatban van, különben hamis</summary>
            public bool isUsed;

            /// <summary>Igaz, ha e-mailt kell küldeni akkor, ha a bit logikai 1-re vált</summary>
            public bool sendEmail;
        }

        /// <summary>Ez az osztály tartalmazza a beállító parancsok leírását</summary>
        public class SetCommand
        {
            /// <summary>Konstruktor</summary>
            /// <param name="numberOfSections">A parancs ennyi adatot tartalmaz</param>
            public SetCommand(byte numberOfSections)
            {
                numOfSections = numberOfSections; // Lementjük az adatok számát,
                sections = new List<SetSection>(numberOfSections); // és inicialiáljuk a lista méretét is
                return;
            }

            /// <summary>A parancs kódja</summary>
            public byte code;

            /// <summary>A parancs neve</summary>
            public string name;

            /// <summary>A parancs adatbájtjainak (DATAn) a száma</summary>
            public byte dataBytes;

            /// <summary>A parancs ennyi adatot tartalmaz</summary>
            public byte numOfSections;

            /// <summary>Igaz, ha a parancsra az eszköz egy standard egybájtos választ küld</summary>
            public bool standardAnswer;

            /// <summary>A parancsra kapott válasz adatainak listája</summary>
            public List<SetSection> sections;
        }

        /// <summary>Ez az osztály tartalmazza a beállító parancsok adatainak leírását</summary>
        public class SetSection
        {
            /// <summary>Konstruktor</summary>
            /// <param name="numberOfBits">Hexadecimális érték esetén a bitek száma</param>
            public SetSection(byte numberOfBits)
            {
                numOfBits = numberOfBits; // Lementjük a bitek számát,
                bits = new List<SetBit>(numberOfBits); // és inicialiáljuk a lista méretét is
                return;
            }

            /// <summary>Az adat neve</summary>
            public string name;

            /// <summary>Amennyiben az adat hexadecimális szám, akkor ez az érték mondja meg, hogy hány bitet tartalmaz</summary>
            public byte numOfBits;

            /// <summary>Az adat típusa</summary>
            public Type type;

            /// <summary>Az adat tizedesjegyeinek a száma</summary>
            /// <remarks>
            /// Ha az érték 0, akkor az adat egész szám.
            /// Ha az érték -1, akkor az adat dátum/idő
            /// Ha az érték -2, akkor az adat hexadecimális szám.
            /// </remarks>
            public sbyte decimals;

            /// <summary>Amennyiben az adat hexadecimális szám, akkor ez a lista tartalmazza a biteket</summary>
            public List<SetBit> bits;
        }

        /// <summary>Ez az osztály tartalmazza a beállító parancsok bitjeinek leírását</summary>
        public class SetBit
        {
            /// <summary>A bit neve</summary>
            public string name;

            /// <summary>Igaz, ha a bit használatban van, különben hamis</summary>
            public bool isUsed;
        }

        /// <summary>A paramétertáblázat elérési útvonala</summary>
        private string filePath;

        /// <summary>Az e-mail címeket tartalmazó táblázat</summary>
        private Object[,] emailListTable;

        /// <summary>Az egyéb Universal Logger beállításokat tartalmazó táblázat</summary>
        private Object[,] settingsTable;

        /// <summary>A lekérdező parancsok adatait tartalmazó táblázat</summary>
        /// <remarks>Nem tartalmazza az általános parancs protokollban meghatározott kötelező parancsokat.</remarks>
        private Object[,] qCommandsTable;

        /// <summary>A lekérdezhető adatok adatait tartalmazó táblázat</summary>
        /// <remarks>Nem tartalmazza az általános parancs protokollban meghatározott kötelező parancsokat.</remarks>
        private Object[,] qSectionsTable;

        /// <summary>A lekérdezhető bitek adatait tartalmazó táblázat</summary>
        /// <remarks>Nem tartalmazza az általános parancs protokollban meghatározott kötelező parancsokat.</remarks>
        private Object[,] qBitsTable;

        /// <summary>A beállító parancsok adatait tartalmazó táblázat</summary>
        /// <remarks>Nem tartalmazza az általános parancs protokollban meghatározott kötelező parancsokat.</remarks>
        private Object[,] sCommandsTable;

        /// <summary>A beállítható adatok adatait tartalmazó táblázat</summary>
        /// <remarks>Nem tartalmazza az általános parancs protokollban meghatározott kötelező parancsokat.</remarks>
        private Object[,] sSectionsTable;

        /// <summary>A beállítható bitek adatait tartalmazó táblázat</summary>
        /// <remarks>Nem tartalmazza az általános parancs protokollban meghatározott kötelező parancsokat.</remarks>
        private Object[,] sBitsTable;

        /// <summary>A program beállításait tartalmazó objektum</summary>
        /// <remarks>Az értékei az Open függvénnyel inicializálhatóak.</remarks>
        public SettingsData settingsData;

        /// <summary>Konstruktor a táblázat elérési útvonalával</summary>
        /// <param name="path">A táblázat elérési útvonala, fájlnévvel együtt</param>
        public OperatingParameters(string path)
        {
            filePath = path; // A táblázat elérési útvonalának beállítása
            settingsData = new SettingsData(); // Inicializáljuk a program beállításait tartalmazó objektumot
        }

        /// <summary>Ha hibás a táblázat, akkor a hiba leírása ebben a változóban lesz</summary>
        public string faultText;

        /// <summary>A függvénnyel beolvasható és ellenőrizhető a paramétertáblázat</summary>
        /// <returns>Igaz, ha érvényesek a táblázat adatai, hamis, ha nem</returns>
        /// <remarks>A függvény nem rendelkezik kivételkezeléssel, így a hívó függvénynek kell ezt megoldania.</remarks>
        public bool Open()
        {
            if (ReadFile()) // Beolvassuk a fájlt, és ha ez sikerült, akkor ellenőrizzük és frissítjük az egyes adatokat,
                return RefreshSettingsData(); // végezetül jelezzük a hívónak, hogy a táblázat helyes volt-e
            return false; // Már a beolvasásnál is hiba történt
        }

        /// <summary>Ez a függvény olvassa be a paramétertáblázatot</summary>
        /// <returns>Igaz, ha érvényesek a táblázat beolvasása sikeres volt, hamis, ha nem</returns>
        /// <remarks>A kivételkezelést nem ez az osztály valósítja meg, ezért ezt felsőbb szinten kell megoldani.</remarks>
        private bool ReadFile()
        {
            FileInfo file = new FileInfo(filePath); // A paramétertáblázat

            using (ExcelPackage package = new ExcelPackage(file/*, "jelszó"*/)) // A paramétertáblázat megnyitása
            { // A using kulcsszó biztosítja, hogy a package objektum törölve legyen a blokk után
                ExcelWorksheet worksheet; // Itt fogjuk tárolni ideiglenesen az egyes munkalapokat

                //
                // Az általános adatokat tartalmazó munkalap beolvasása
                //
                worksheet = package.Workbook.Worksheets[1]; // A munkalap,
                emailListTable = (System.Object[,])worksheet.Cells[2, 1, 11, 1].Value; // az A2 - A11 értékek,
                settingsTable = (System.Object[,])worksheet.Cells[1, 3, 14, 3].Value; // és a C1 - C14 értékek beolvasása

                try
                {
                    settingsData.numOfQueryCommands = Convert.ToByte(settingsTable[7, 0]); // A lekérdező parancsok számának,
                    settingsData.numOfQuerySections = Convert.ToUInt16(settingsTable[9, 0]); // a lekérdezhető adatok számának,
                    settingsData.numOfQueryBits = Convert.ToInt32(settingsTable[10, 0]); // és a lekérdezhető bitek számának a beállítása

                    if (settingsData.numOfQueryCommands < 3 || // A rendszerben található lekérdező parancsok száma minimum 3,
                        31 < settingsData.numOfQueryCommands) // maximum 31 lehet (a kötelező parancsokat is számba véve)
                    {
                        faultText = "The number of the query commands in the C8 cell should minimum 3, and maximum 31!";
                        return false; // Ha ez nem teljesül, akkor jelezzük, hogy hibás a táblázat
                    }

                    settingsData.numOfSetCommands = Convert.ToByte(settingsTable[8, 0]); // A beállító parancsok számának,
                    settingsData.numOfSetSections = Convert.ToUInt16(settingsTable[11, 0]); // a beállítható adatok számának,
                    settingsData.numOfSetBits = Convert.ToInt32(settingsTable[12, 0]); // és a beállítható bitek számának a beállítása

                    if (settingsData.numOfSetCommands < 2 || // A rendszerben található beállító parancsok számánal minimuma 2, maximuma
                        (256 - settingsData.numOfQueryCommands) < settingsData.numOfSetCommands) // pedig a beállító parancsok számától
                    { // függ
                        faultText = "The number of the set commands in the C9 cell should minimum 2, and maximum " +
                            (256 - settingsData.numOfQueryCommands).ToString() + "!";
                        return false; // Ha a feltételek nem teljesülnek, akkor jelezzük, hogy hibás a táblázat
                    }
                }
                catch // Ha valamelyik szám kívül esik a megengedett tartományon,
                {
                    faultText = "One of the numbers in the C8 - C13 cell is incorrect!" + Environment.NewLine +
                        "The allowed ranges of the different numbers are:" + Environment.NewLine +
                        "0 <= number of query commands <= 255" + Environment.NewLine +
                        "0 <= number of query sections <= 65535" + Environment.NewLine +
                        "0 <= number of query bits <= 4294967295" + Environment.NewLine +
                        "0 <= number of set commands <= 255" + Environment.NewLine +
                        "0 <= number of set sections <= 65535" + Environment.NewLine +
                        "0 <= number of set bits <= 4294967295";
                    return false; // akkor jelezzük, hogy hibás a táblázat
                }

                int qCommandsEnd = (int)(settingsData.numOfQueryCommands + 1); // A lekérdező parancsokhoz,
                int qSectionsEnd = (int)(settingsData.numOfQuerySections + 1); // a lekérdezhető adatokhoz,
                int qBitsEnd = settingsData.numOfQueryBits + 1; // a lekérdezhető bitekhez,
                int sCommandsEnd = (int)(settingsData.numOfSetCommands + 1); // a beállító parancsokhoz,
                int sSectionsEnd = (int)(settingsData.numOfSetSections + 1); // a beállítható adatokhoz,
                int sBitsEnd = settingsData.numOfSetBits + 1; // és a beállítható bitekhez tartozó táblázat utolsó sorának száma

                settingsData.numOfQueryCommands -= 3; // Az általános parancs protokollban lévő kötelező parancsokat,
                settingsData.numOfSetCommands -= 2;
                settingsData.numOfQuerySections -= 2; // illetve kötelező adatokat nem számoljuk bele az értékekbe
                settingsData.numOfSetSections -= 1;

                settingsData.queryCommands = new List<QueryCommand>(settingsData.numOfQueryCommands); // Létrehozzuk a parancsokat tároló
                settingsData.setCommands = new List<SetCommand>(settingsData.numOfSetCommands); // listákat

                //
                // A lekérdező parancsok adatait tartalmazó munkalap beolvasása
                //
                worksheet = package.Workbook.Worksheets[2]; // A munkalap,
                if (5 <= qCommandsEnd) // a lekérdező parancsok (A5 - E?),
                    qCommandsTable = (System.Object[,])worksheet.Cells[5, 1, qCommandsEnd, 5].Value;
                if (4 <= qSectionsEnd) // a lekérdezhető adatok (G4 - K?),
                    qSectionsTable = (System.Object[,])worksheet.Cells[4, 7, qSectionsEnd, 11].Value;
                if (2 <= qBitsEnd) // és a lekérdezhető bitek (M2 - O?) beolvasása
                    qBitsTable = (System.Object[,])worksheet.Cells[2, 13, qBitsEnd, 15].Value;

                //
                // A beállító parancsok adatait tartalmazó munkalap beolvasása
                //
                worksheet = package.Workbook.Worksheets[3]; // A munkalap,
                if (4 <= sCommandsEnd) // a beállító parancsok (A4 - E?),
                    sCommandsTable = (System.Object[,])worksheet.Cells[4, 1, sCommandsEnd, 5].Value;
                if (3 <= sSectionsEnd) // a beállítható adatok (G3 - J?),
                    sSectionsTable = (System.Object[,])worksheet.Cells[3, 7, sSectionsEnd, 10].Value;
                if (2 <= sBitsEnd) // és a beállítható bitek (L2 - M?) beolvasása
                    sBitsTable = (System.Object[,])worksheet.Cells[2, 12, sBitsEnd, 13].Value;
            }

            return true; // A táblázat beolvasása rendben végbement
        }

        /// <summary>Ez a függvény ellenőrzi, hogy érvényesek-e a beolvasott táblázat adatai</summary>
        /// <returns>Igaz, ha érvényesek a táblázat adatai, hamis, ha nem</returns>
        private bool RefreshSettingsData()
        {
            for (byte i = 0; i < 10; i++) // Végigmegyünk az összes e-mail címen
            {
                string tmp = Convert.ToString(emailListTable[i, 0]); // Beolvassuk az adott sort
                if (tmp != "") // Ha nem üres a sor,
                {
                    if (ValidateMail(tmp)) // és szintaktikailag helyes az e-mail cím
                    {
                        settingsData.emailList.Add(tmp); // akkor hozzáadjuk az e-mail címet a listához
                    }
                    else
                    {
                        faultText = "The e-mail address in the A" + (i + 2).ToString() + " cell is invalid!";
                        return false; // Ha az e-mail cím helytelen, akkor jelezzük, hogy hibás a táblázat
                    }
                }
            }

            settingsData.port = Convert.ToString(settingsTable[0, 0]); // A használt COM port nevének beolvasása
            try
            {
                settingsData.baudRate = Convert.ToInt32(settingsTable[1, 0]); // A baud rate beolvasása
                if (settingsData.baudRate != 600 && settingsData.baudRate != 1200 && // Ha nincs az érték az
                    settingsData.baudRate != 2400 && settingsData.baudRate != 4800 && // engedélyezett listában,
                    settingsData.baudRate != 9600 && settingsData.baudRate != 19200 &&
                    settingsData.baudRate != 38400 && settingsData.baudRate != 57600 &&
                    settingsData.baudRate != 115200 && settingsData.baudRate != 230400)
                {
                    faultText = "The baud rate in the C2 cell is not in the allowed list!";
                    return false; // akkor jelezzük, hogy hibás a táblázat
                }
            }
            catch // Ha a szám nem 32 bites, előjeles egész,
            {
                faultText = "The baud rate in the C2 cell is not in the allowed list!";
                return false; // akkor jelezzük, hogy hibás a táblázat
            }

            try
            {
                settingsData.timeout = Convert.ToUInt16(settingsTable[2, 0]); // A kommunikáció során használt időtúllépés beolvasása
            }
            catch // Ha a szám nem előjel nélküli 16 bites egész,
            {
                faultText = "The communication timeout in the C3 cell should be minimum 0s, and maximum 65535s!";
                return false; // akkor jelezzük, hogy hibás a táblázat
            }

            try
            {
                settingsData.systemId = Convert.ToUInt16(settingsTable[3, 0]); // A rendszer azonosítószámának a beolvasása
            }
            catch // Ha a szám nem előjel nélküli 16 bites egész,
            {
                faultText = "The system ID in the C4 cell should be minimum 0, and maximum 65535!";
                return false; // akkor jelezzük, hogy hibás a táblázat
            }

            settingsData.systemName = Convert.ToString(settingsTable[4, 0]); // A rendszer nevének a beolvasása
            if (settingsData.systemName.Length < 1 || 64 < settingsData.systemName.Length) // Ha túl rövid, vagy túl hosszú a név
            { // (minimum 1, maximum 64 karakter),
                faultText = "The length of the system's name in the C5 cell should be minimum 1 character, and maximum 64 characters!";
                return false; // akkor jelezzük, hogy hibás a táblázat
            }

            try
            {
                settingsData.emailSendingHour = Convert.ToUInt16(settingsTable[5, 0]); // Az e-mail küldési idő beolvasása
                if (23 < settingsData.emailSendingHour) // Ha nem értelmezhető az óra,
                {
                    faultText = "The hour of the e-mail sending in the C6 cell should be maximum 23!";
                    return false; // akkor jelezzük, hogy hibás a táblázat
                }
            }
            catch // Ha a szám nem 16 bites, előjel nélkül egész,
            {
                faultText = "The hour of the e-mail sending in the C6 cell should be minimum 0, and maximum 65535!";
                return false; // akkor jelezzük, hogy hibás a táblázat
            }

            //
            // A logfájlok könyvtárának a neve legalább 1 karakter kell legyen, de nem lehet több 32 karakternél,
            // csak angol betűket (A-Z, a-z), számokat (0-9), kötőjelet vagy alulvonást tartalmazhat, és érvényes
            // névnek kell lennie (azaz nem lehet PRN, prn, AUX, aux, CLOCK, clock, NUL, nul, CON, con, COMx, comx,
            // LPTx, lptx (ahol x egy számot takar))
            //
            settingsData.logDirectoryName = Convert.ToString(settingsTable[6, 0]); // A könyvtárnév beolvasása
            if (settingsData.logDirectoryName.Length < 1 || 32 < settingsData.logDirectoryName.Length) // Ha túl rövid, vagy túl hosszú
            { // a név,
                faultText = "The length of the directory name in the C7 cell should be minimum 1 character, and maximum 32 characters!";
                return false; // akkor jelezzük, hogy hibás a táblázat
            }
            string pattern1 = @"^(PRN|prn|AUX|aux|CLOCK\$|clock\$|NUL|nul|CON|con|COM\d|com\d|LPT\d|lpt\d)$";
            string pattern2 = @"^([A-Z]|[a-z]|[0-9]|_|-){1,32}$";
            if (Regex.IsMatch(settingsData.logDirectoryName, pattern1) || // Ha a könyvtár neve nem érvényes,
                !Regex.IsMatch(settingsData.logDirectoryName, pattern2))
            { // a név,
                faultText = "The name of the directory in the C7 cell contains invalid characters!" + Environment.NewLine +
                    "It should only contains english letters (A-Z, a-z), numbers (0-9)," + Environment.NewLine +
                    "hyphen (-) or underscore(_), and should be a valid directory name," + Environment.NewLine +
                    "so it can't be PRN, prn, AUX, aux, CLOCK, clock, NUL, nul, CON, con," + Environment.NewLine +
                    "COMx, comx, LPTx or lptx (where x refers to a number).";
                return false; // akkor jelezzük, hogy hibás a táblázat
            }

            string timerAutoStartString = Convert.ToString(settingsTable[13, 0]); // Az időzítő automatikus indításához tartozó szöveg beolvasása
            try
            {
                switch (timerAutoStartString) // A szöveg alapján kiválasztjuk a megfelelő számot
                {
                    case "Don't start timer": settingsData.timerAutoStart = 0; break;
                    case "1s": settingsData.timerAutoStart = 1; break;
                    case "2s": settingsData.timerAutoStart = 2; break;
                    case "5s": settingsData.timerAutoStart = 3; break;
                    case "10s": settingsData.timerAutoStart = 4; break;
                    case "20s": settingsData.timerAutoStart = 5; break;
                    case "30s": settingsData.timerAutoStart = 6; break;
                    case "1min": settingsData.timerAutoStart = 7; break;
                    case "2min": settingsData.timerAutoStart = 8; break;
                    case "5min": settingsData.timerAutoStart = 9; break;
                    case "10min": settingsData.timerAutoStart = 10; break;
                    case "30min": settingsData.timerAutoStart = 11; break;
                    case "60min": settingsData.timerAutoStart = 12; break;

                    default: // A cella értéke más, mint a listában lévő értékek
                        settingsData.timerAutoStart = 0;
                        faultText = "The timer autostart interval in the C14 cell is not in the allowed list!";
                        return false; // akkor jelezzük, hogy hibás a táblázat
                }
            }
            catch // Ha a szám nem 32 bites, előjeles egész,
            {
                faultText = "The timer autostart interval in the C14 cell is not in the allowed list!";
                return false; // akkor jelezzük, hogy hibás a táblázat
            }

            // A beállító és lekérdező parancsok számát nem ellenőrizzük, mivel ez már megtörtént a ReadFile() függvényben

            ushort sectionCounter = 0; // Itt tároljuk, hogy mennyit olvastunk ki az adatokból,
            int bitCounter = 0; // illetve a bitekből
            for (byte i = 0; i < settingsData.numOfQueryCommands; i++) // Végigmegyünk a lekérdező parancsokon
            {
                byte codeTmp;
                try
                {
                    codeTmp = Convert.ToByte(Convert.ToString(qCommandsTable[i, 0]), 16); // A parancs kódjának beolvasása

                    if (codeTmp < 0x04 || (0x1F < codeTmp && codeTmp < 0xA0)) // Ha a parancs kódja kívül esik a megengedett tartomnányon,
                    {
                        faultText = "The code of the query command in the A" + (i + 5).ToString() + " cell is invalid!" +
                            Environment.NewLine + "It should be in the range of 0x04 - 0x1F, or 0xA0 - 0xFF!";
                        return false; // akkor jelezzük, hogy hibás a táblázat
                    }

                    for (byte j = 0; j < settingsData.queryCommands.Count; j++) // Végigmegyünk a már meglévő parancsokon,
                    {
                        if (settingsData.queryCommands[j].code == codeTmp) // Ha a parancs kódja szerepelt már korábban (azaz a kód
                        { // duplikálva van),
                            faultText = "The code of the query command in the A" + (i + 5).ToString() + " cell is invalid!" +
                                Environment.NewLine + "It is the same as the query command's code in the A" + (j + 5).ToString() + " cell!";
                            return false; // akkor jelezzük, hogy hibás a táblázat
                        }
                    }
                }
                catch // Ha a parancs kódja nem előjel nélküli 8 bites egész,
                {
                    faultText = "The code of the query command in the A" + (i + 5).ToString() + " cell is invalid!" + Environment.NewLine +
                        Environment.NewLine + "It should be in the range of 0x00 - 0xFF!";
                    return false; // akkor jelezzük, hogy hibás a táblázat
                }

                string nameTmp = Convert.ToString(qCommandsTable[i, 1]); // A parancs nevének a beolvasása
                if (nameTmp.Length < 1 || 32 < nameTmp.Length) // Ha túl rövid, vagy túl hosszú a név (minimum 1, maximum 32 karakter),
                {
                    faultText = "The length of the query command's name in the B" + (i + 5).ToString() + Environment.NewLine +
                        " cell should be minimum 1 character, and maximum 32 characters!";
                    return false; // akkor jelezzük, hogy hibás a táblázat
                }

                byte bytesTmp;
                try
                {
                    bytesTmp = Convert.ToByte(qCommandsTable[i, 2]); // A parancsra kapott válasz adatbájtainak (DATAn) számának a beolvasása

                    if (bytesTmp == 0) // Ha az érték érvénytelen,
                    {
                        faultText = "The number of the query command's data bytes in the C" + (i + 5).ToString() + Environment.NewLine +
                            " cell should be greater, than 0!";
                        return false; // akkor jelezzük, hogy hibás a táblázat
                    }
                }
                catch // Ha a válasz adatbájtjainak a száma nem előjel nélküli 8 bites egész,
                {
                    faultText = "The number of the query command's data bytes in the C" + (i + 5).ToString() + Environment.NewLine +
                        " cell should be minimum 0, and maximum 255!";
                    return false; // akkor jelezzük, hogy hibás a táblázat
                }

                byte sectionsTmp;
                try
                {
                    sectionsTmp = Convert.ToByte(qCommandsTable[i, 3]); // A parancsra kapott válaszban lévő adatok számának a beolvasása

                    if (sectionsTmp == 0 || bytesTmp < sectionsTmp) // Ha az érték érvénytelen,
                    {
                        faultText = "The number of the query command's sections in the D" + (i + 5).ToString() +
                            " cell should be greater, than 0," + Environment.NewLine +
                            "and not greater, than the number of this command's data bytes in the C" + (i + 5).ToString() + " cell!";
                        return false; // akkor jelezzük, hogy hibás a táblázat
                    }
                }
                catch // Ha a válaszban lévő adatok száma nem előjel nélküli 8 bites egész,
                {
                    faultText = "The number of the query command's sections in the D" + (i + 5).ToString() + Environment.NewLine +
                        "cell should be minimum 0, and maximum 255!";
                    return false; // akkor jelezzük, hogy hibás a táblázat
                }

                bool autoQueryTmp;
                try
                {
                    autoQueryTmp = Convert.ToBoolean(qCommandsTable[i, 4]); // Annak beolvasása, hogy be kell-e pipálni a parancs melletti jelölőnégyzetet
                }
                catch // Ha a beolvasott adat nem logikai érték,
                {
                    faultText = "The auto query value of the query command in the E" + (i + 5).ToString() + Environment.NewLine +
                        "cell should be true or false!";
                    return false; // akkor jelezzük, hogy hibás a táblázat
                }

                // A vizsgált sor elemei rendben vannak
                settingsData.queryCommands.Add(new QueryCommand(sectionsTmp)); // Létrehozzuk az új parancsot,
                settingsData.queryCommands[i].code = codeTmp; // majd lementjük az adatait
                settingsData.queryCommands[i].name = nameTmp;
                settingsData.queryCommands[i].dataBytes = bytesTmp;
                settingsData.queryCommands[i].autoQuery = autoQueryTmp;

                for (ushort j = sectionCounter; j < sectionCounter + sectionsTmp; j++) // Végigmegyünk a parancshoz tartozó adatokon
                {
                    nameTmp = Convert.ToString(qSectionsTable[j, 0]); // Az adat nevének a beolvasása
                    if (nameTmp.Length < 1 || 32 < nameTmp.Length) // Ha túl rövid, vagy túl hosszú a név (minimum 1, maximum 32 karakter),
                    {
                        faultText = "The length of the query section's name in the F" + (j + 4).ToString() + Environment.NewLine +
                            "cell should be minimum 1 character, and maximum 32 characters!";
                        return false; // akkor jelezzük, hogy hibás a táblázat
                    }
                    
                    try
                    {
                        bytesTmp = Convert.ToByte(qSectionsTable[j, 1]); // A adat méretének a beolvasása

                        if (bytesTmp != 1 && bytesTmp != 2 && bytesTmp != 4 && bytesTmp != 7) // Ha az érték érvénytelen,
                        {
                            faultText = "The size of the query section in the G" + (j + 4).ToString() + Environment.NewLine +
                                "cell should be 1, 2, 4 or 7 bytes!";
                            return false; // akkor jelezzük, hogy hibás a táblázat
                        }
                    }
                    catch // Ha az adat mérete nem előjel nélküli 8 bites egész,
                    {
                        faultText = "The size of the query section in the G" + (j + 4).ToString() + Environment.NewLine +
                            "cell should be minimum 0, and maximum 255 bytes!";
                        return false; // akkor jelezzük, hogy hibás a táblázat
                    }

                    bool signedTmp;
                    try
                    {
                        signedTmp = Convert.ToBoolean(qSectionsTable[j, 2]); // Annak beolvasása, hogy az adat előjeles-e
                    }
                    catch // Ha a beolvasott adat nem logikai érték,
                    {
                        faultText = "The signedness of the query section in the H" + (j + 4).ToString() + Environment.NewLine +
                            "cell should be true or false!";
                        return false; // akkor jelezzük, hogy hibás a táblázat
                    }

                    sbyte decimalsTmp;
                    byte bitsTmp;
                    try
                    {
                        decimalsTmp = Convert.ToSByte(qSectionsTable[j, 3]); // Az adat tizedesjegyeinek a számának a beolvasása
                        if (decimalsTmp < -2 || 4 < decimalsTmp) // Ha az érték érvénytelen,
                        {
                            faultText = "The number of the query section's decimals in the I" + (j + 4).ToString() + Environment.NewLine +
                                "cell should be minimum -2, and maximum 4!";
                            return false; // akkor jelezzük, hogy hibás a táblázat
                        }
                        if(decimalsTmp == -2) bitsTmp = (byte)(bytesTmp * 8);
                        else bitsTmp = 0;

                        if (decimalsTmp == -2 && (signedTmp || bytesTmp == 7)) // Ha az adat hexadecimális, azonban előjeles, vagy pedig
                        { // a változó 7 bájtos,
                            faultText = "The number of the query section's decimals in the I" + (j + 4).ToString() + Environment.NewLine +
                                "cell shouldn't be -2, as the section is signed and/or its size is 7 bytes!";
                            return false; // akkor jelezzük, hogy hibás a táblázat
                        }

                        if (decimalsTmp == -1 && (signedTmp || bytesTmp != 7)) // Ha az adat dátum/idő, azonban előjeles, vagy pedig
                        { // a változó nem 7 bájtos,
                            faultText = "The number of the query section's decimals in the I" + (j + 4).ToString() + Environment.NewLine +
                                "cell shouldn't be -1, as the section is signed and/or its size isn't 7 bytes!";
                            return false; // akkor jelezzük, hogy hibás a táblázat
                        }

                        if (decimalsTmp != -1 && bytesTmp == 7) // Ha az adat nem dátum/idő, azonban a változó 7 bájtos,
                        {
                            faultText = "The number of the query section's decimals in the I" + (j + 4).ToString() + Environment.NewLine +
                                "cell should be -1, as the section's size is 7 bytes!";
                            return false; // akkor jelezzük, hogy hibás a táblázat
                        }
                    }
                    catch // Ha az adat tizedesjegyeinek a száma nem előjel nélküli 8 bites egész,
                    {
                        faultText = "The number of the query section's decimals in the I" + (j + 4).ToString() + Environment.NewLine +
                            "cell should be minimum 0, and maximum 255!";
                        return false; // akkor jelezzük, hogy hibás a táblázat
                    }

                    bool saveTmp;
                    try
                    {
                        saveTmp = Convert.ToBoolean(qSectionsTable[j, 4]); // Az adatmentés szükségességének a beolvasása
                    }
                    catch // Ha az adatmentés szükségessége nem logikai érték,
                    {
                        faultText = "The save value of the query section in the J" + (j + 4).ToString() + Environment.NewLine +
                            "cell should be true or false!";
                        return false; // akkor jelezzük, hogy hibás a táblázat
                    }

                    settingsData.queryCommands[i].sections.Add(new QuerySection(bitsTmp)); // Létrehozzuk az új adatot,
                    settingsData.queryCommands[i].sections[j - sectionCounter].name = nameTmp; // majd lementjük az adatait
                    settingsData.queryCommands[i].sections[j - sectionCounter].decimals = decimalsTmp;
                    settingsData.queryCommands[i].sections[j - sectionCounter].save = saveTmp;
                    
                    if (decimalsTmp == -1) // Ha az adat dátum/idő,
                    {
                        if (saveTmp) // és le is kell ezt menteni
                            settingsData.useDeviceDateTime = true; // akkor jelezzük, hogy nem a pc idejét használjuk a mentésekhez

                        // Lementjük az adat típusát
                        settingsData.queryCommands[i].sections[j - sectionCounter].type = typeof(DateTime);
                    }
                    else // Ha az adat nem dátum/idő, akkor megvizsgáljuk, hogy milyen formátumú
                    {
                        if(signedTmp) // Előjeles
                        {
                            if(bytesTmp == 1) // 8 bites előjeles egész
                                settingsData.queryCommands[i].sections[j - sectionCounter].type = typeof(sbyte);
                            else if(bytesTmp == 2) // 16 bites előjeles egész
                                settingsData.queryCommands[i].sections[j - sectionCounter].type = typeof(short);
                            else // 32 bites előjeles egész
                                settingsData.queryCommands[i].sections[j - sectionCounter].type = typeof(int);
                        }
                        else // Előjel nélküli
                        {
                            if(bytesTmp == 1) // 8 bites előjel nélküli egész
                                settingsData.queryCommands[i].sections[j - sectionCounter].type = typeof(byte);
                            else if(bytesTmp == 2) // 16 bites előjel nélküli egész
                                settingsData.queryCommands[i].sections[j - sectionCounter].type = typeof(ushort);
                            else // 32 bites előjel nélküli egész
                                settingsData.queryCommands[i].sections[j - sectionCounter].type = typeof(uint);
                        }
                    }
                    
                    for (int k = bitCounter; k < bitCounter + bitsTmp; k++) // Végigmegyünk az adathoz tartozó biteken
                    {
                        nameTmp = Convert.ToString(qBitsTable[k, 0]); // Az adat nevének a beolvasása
                        if (nameTmp.Length < 1 || 32 < nameTmp.Length) // Ha túl rövid, vagy túl hosszú a név (minimum 1, maximum
                        { // 32 karakter),
                            faultText = "The length of the query bits's name in the L" + (k + 2).ToString() + Environment.NewLine +
                                "cell should be minimum 1 character, and maximum 32 characters!";
                            return false; // akkor jelezzük, hogy hibás a táblázat
                        }

                        bool isUsedTmp;
                        try
                        {
                            isUsedTmp = Convert.ToBoolean(qBitsTable[k, 1]); // Annak beolvasása, hogy a bit használatban van-e
                        }
                        catch // Ha a beolvasott adat nem logikai érték,
                        {
                            faultText = "The 'is used' value of the query section in the M" + (k + 2).ToString() + Environment.NewLine +
                                "cell should be true or false!";
                            return false; // akkor jelezzük, hogy hibás a táblázat
                        }

                        bool sendEmailTmp;
                        try
                        { // Annak beolvasása, hogy kell-e e-mailt küldeni,
                            sendEmailTmp = Convert.ToBoolean(qBitsTable[k, 2]); // ha a bit logikai 1-re vált
                        }
                        catch // Ha a beolvasott adat nem logikai érték,
                        {
                            faultText = "The 'send mail' value of the query section in the N" + (k + 2).ToString() + Environment.NewLine +
                                "cell should be true or false!";
                            return false; // akkor jelezzük, hogy hibás a táblázat
                        }

                        settingsData.queryCommands[i].sections[j - sectionCounter].bits.Add(new QueryBit()); // Létrehozzuk az új bitet,
                        settingsData.queryCommands[i].sections[j - sectionCounter].bits[k - bitCounter].name = nameTmp; // majd lementjük
                        settingsData.queryCommands[i].sections[j - sectionCounter].bits[k - bitCounter].isUsed = isUsedTmp; // az adatait
                        settingsData.queryCommands[i].sections[j - sectionCounter].bits[k - bitCounter].sendEmail = sendEmailTmp;
                    }
                    
                    bitCounter += bitsTmp; // A bitek számlálóját eltoljuk a már kiolvasott bitek számával
                }

                sectionCounter += sectionsTmp; // Az adatok számlálóját eltoljuk a már kiolvasott adatok számával
            }

            sectionCounter = 0; // Itt tároljuk, hogy mennyit olvastunk ki az adatokból,
            bitCounter = 0; // illetve a bitekből
            for (byte i = 0; i < settingsData.numOfSetCommands; i++) // Végigmegyünk a beállító parancsokon
            {
                byte codeTmp;
                try
                {
                    codeTmp = Convert.ToByte(Convert.ToString(sCommandsTable[i, 0]), 16); // A parancs kódjának beolvasása

                    if (codeTmp < 0x04 || codeTmp == 0x20) // Ha a parancs kódja kívül esik a megengedett tartomnányon,
                    {
                        faultText = "The code of the set command in the A" + (i + 4).ToString() + " cell is invalid!" +
                            Environment.NewLine + "It should be in the range of 0x04 - 0x1F, or 0x21 - 0xFF!";
                        return false; // akkor jelezzük, hogy hibás a táblázat
                    }

                    for (byte j = 0; j < settingsData.queryCommands.Count; j++) // Végigmegyünk a már meglévő lekérdező parancsokon,
                    {
                        if (settingsData.queryCommands[j].code == codeTmp) // Ha a parancs kódja szerepelt már korábban (azaz a kód
                        { // duplikálva van),
                            faultText = "The code of the set command in the A" + (i + 4).ToString() + " cell is invalid!" +
                                Environment.NewLine + "It is the same as the query command's code in the A" + (j + 5).ToString() + " cell!";
                            return false; // akkor jelezzük, hogy hibás a táblázat
                        }
                    }

                    for (byte j = 0; j < settingsData.setCommands.Count; j++) // Végigmegyünk a már meglévő beállító parancsokon,
                    {
                        if (settingsData.setCommands[j].code == codeTmp) // Ha a parancs kódja szerepelt már korábban (azaz a kód
                        { // duplikálva van),
                            faultText = "The code of the set command in the A" + (i + 4).ToString() + " cell is invalid!" +
                                Environment.NewLine + "It is the same as the set command's code in the A" + (j + 4).ToString() + " cell!";
                            return false; // akkor jelezzük, hogy hibás a táblázat
                        }
                    }
                }
                catch // Ha a parancs kódja nem előjel nélküli 8 bites egész,
                {
                    faultText = "The code of the set command in the A" + (i + 4).ToString() + " cell is invalid!" + Environment.NewLine +
                        Environment.NewLine + "It should be in the range of 0x00 - 0xFF!";
                    return false; // akkor jelezzük, hogy hibás a táblázat
                }

                string nameTmp = Convert.ToString(sCommandsTable[i, 1]); // A parancs nevének a beolvasása
                if (nameTmp.Length < 1 || 32 < nameTmp.Length) // Ha túl rövid, vagy túl hosszú a név (minimum 1, maximum 32 karakter),
                {
                    faultText = "The length of the set command's name in the B" + (i + 4).ToString() + Environment.NewLine +
                        "cell should be minimum 1 character, and maximum 32 characters!";
                    return false; // akkor jelezzük, hogy hibás a táblázat
                }

                byte bytesTmp;
                try
                {
                    bytesTmp = Convert.ToByte(sCommandsTable[i, 2]); // A parancs adatbájtainak (DATAn) a számának a beolvasása
                }
                catch // Ha a parancs adatbájtjainak a száma nem előjel nélküli 8 bites egész,
                {
                    faultText = "The number of the set command's data bytes in the C" + (i + 4).ToString() + Environment.NewLine +
                        "cell should be minimum 0, and maximum 255!";
                    return false; // akkor jelezzük, hogy hibás a táblázat
                }

                byte sectionsTmp;
                try
                {
                    sectionsTmp = Convert.ToByte(sCommandsTable[i, 3]); // A parancsban lévő adatok számának a beolvasása

                    if (bytesTmp < sectionsTmp) // Ha az érték érvénytelen,
                    {
                        faultText = "The number of the set command's sections in the D" + (i + 4).ToString() + Environment.NewLine +
                            "cell should be not greater, than the number of this command's" + Environment.NewLine +
                            "data bytes in the C" + (i + 4).ToString() + " cell!";
                        return false; // akkor jelezzük, hogy hibás a táblázat
                    }
                }
                catch // Ha a parancsban lévő adatok száma nem előjel nélküli 8 bites egész,
                {
                    faultText = "The number of the set command's sections in the D" + (i + 4).ToString() + Environment.NewLine +
                        "cell should be minimum 0, and maximum 255!";
                    return false; // akkor jelezzük, hogy hibás a táblázat
                }

                bool answerTmp;
                try
                {
                    answerTmp = Convert.ToBoolean(sCommandsTable[i, 4]); // Annak beolvasása, hogy kapunk-e egybájtos választ a parancsra
                }
                catch // Ha a beolvasott adat nem logikai érték,
                {
                    faultText = "The answer of the set section in the E" + (i + 4).ToString() + Environment.NewLine +
                        "cell should be true or false!";
                    return false; // akkor jelezzük, hogy hibás a táblázat
                }

                // A vizsgált sor elemei rendben vannak

                settingsData.setCommands.Add(new SetCommand(sectionsTmp)); // Létrehozzuk az új parancsot,
                settingsData.setCommands[i].code = codeTmp; // majd lementjük az adatait
                settingsData.setCommands[i].name = nameTmp;
                settingsData.setCommands[i].dataBytes = bytesTmp;
                settingsData.setCommands[i].standardAnswer = answerTmp;

                for (ushort j = sectionCounter; j < sectionCounter + sectionsTmp; j++) // Végigmegyünk a parancshoz tartozó adatokon
                {
                    nameTmp = Convert.ToString(sSectionsTable[j, 0]); // Az adat nevének a beolvasása
                    if (nameTmp.Length < 1 || 32 < nameTmp.Length) // Ha túl rövid, vagy túl hosszú a név (minimum 1, maximum 32 karakter),
                    {
                        faultText = "The length of the set section's name in the G" + (j + 3).ToString() + Environment.NewLine +
                            "cell should be minimum 1 character, and maximum 32 characters!";
                        return false; // akkor jelezzük, hogy hibás a táblázat
                    }
                    
                    try
                    {
                        bytesTmp = Convert.ToByte(sSectionsTable[j, 1]); // A adat méretének a beolvasása

                        if (bytesTmp != 1 && bytesTmp != 2 && bytesTmp != 4 && bytesTmp != 7) // Ha az érték érvénytelen,
                        {
                            faultText = "The size of the set section in the H" + (j + 3).ToString() + Environment.NewLine +
                                "cell should be 1, 2, 4 or 7 bytes!";
                            return false; // akkor jelezzük, hogy hibás a táblázat
                        }
                    }
                    catch // Ha az adat mérete nem előjel nélküli 8 bites egész,
                    {
                        faultText = "The size of the set section in the H" + (j + 3).ToString() + Environment.NewLine +
                            "cell should be minimum 0, and maximum 255 bytes!";
                        return false; // akkor jelezzük, hogy hibás a táblázat
                    }

                    bool signedTmp;
                    try
                    {
                        signedTmp = Convert.ToBoolean(sSectionsTable[j, 2]); // Annak beolvasása, hogy az adat előjeles-e
                    }
                    catch // Ha a beolvasott adat nem logikai érték,
                    {
                        faultText = "The signedness of the set section in the I" + (j + 3).ToString() + Environment.NewLine +
                            "cell should be true or false!";
                        return false; // akkor jelezzük, hogy hibás a táblázat
                    }

                    sbyte decimalsTmp;
                    byte bitsTmp;
                    try
                    {
                        decimalsTmp = Convert.ToSByte(sSectionsTable[j, 3]); // Az adat tizedesjegyeinek a számának a beolvasása
                        if (decimalsTmp < -2 || 4 < decimalsTmp) // Ha az érték érvénytelen,
                        {
                            faultText = "The number of the set section's decimals in the J" + (j + 3).ToString() + Environment.NewLine +
                                "cell should be minimum -2, and maximum 4!";
                            return false; // akkor jelezzük, hogy hibás a táblázat
                        }
                        if(decimalsTmp == -2) bitsTmp = (byte)(bytesTmp * 8);
                        else bitsTmp = 0;

                        if (decimalsTmp == -2 && (signedTmp || bytesTmp == 7)) // Ha az adat hexadecimális, azonban előjeles, vagy pedig
                        { // a változó 7 bájtos,
                            faultText = "The number of the set section's decimals in the J" + (j + 3).ToString() + Environment.NewLine +
                                "cell shouldn't be -2, as the section is signed and/or its size is 7 bytes!";
                            return false; // akkor jelezzük, hogy hibás a táblázat
                        }

                        if (decimalsTmp == -1 && (signedTmp || bytesTmp != 7)) // Ha az adat dátum/idő, azonban előjeles, vagy pedig
                        { // a változó nem 7 bájtos,
                            faultText = "The number of the set section's decimals in the J" + (j + 3).ToString() + Environment.NewLine +
                                "cell shouldn't be -1, as the section is signed and/or its size isn't 7 bytes!";
                            return false; // akkor jelezzük, hogy hibás a táblázat
                        }

                        if (decimalsTmp != -1 && bytesTmp == 7) // Ha az adat nem dátum/idő, azonban a változó 7 bájtos,
                        {
                            faultText = "The number of the set section's decimals in the J" + (j + 3).ToString() + Environment.NewLine +
                                "cell should be -1, as the section's size is 7 bytes!";
                            return false; // akkor jelezzük, hogy hibás a táblázat
                        }
                    }
                    catch // Ha az adat tizedesjegyeinek a száma nem előjel nélküli 8 bites egész,
                    {
                        faultText = "The number of the set section's decimals in the J" + (j + 3).ToString() + Environment.NewLine +
                            "cell should be minimum 0, and maximum 255!";
                        return false; // akkor jelezzük, hogy hibás a táblázat
                    }

                    settingsData.setCommands[i].sections.Add(new SetSection(bitsTmp)); // Létrehozzuk az új adatot,
                    settingsData.setCommands[i].sections[j - sectionCounter].name = nameTmp; // majd lementjük az adatait
                    settingsData.setCommands[i].sections[j - sectionCounter].decimals = decimalsTmp;
                    
                    if (decimalsTmp == -1) // Ha az adat dátum/idő,
                    { // akkor lementjük az adat típusát
                        settingsData.setCommands[i].sections[j - sectionCounter].type = typeof(DateTime);
                    }
                    else // Ha az adat nem dátum/idő, akkor megvizsgáljuk, hogy milyen formátumú
                    {
                        if(signedTmp) // Előjeles
                        {
                            if(bytesTmp == 1) // 8 bites előjeles egész
                                settingsData.setCommands[i].sections[j - sectionCounter].type = typeof(sbyte);
                            else if(bytesTmp == 2) // 16 bites előjeles egész
                                settingsData.setCommands[i].sections[j - sectionCounter].type = typeof(short);
                            else // 32 bites előjeles egész
                                settingsData.setCommands[i].sections[j - sectionCounter].type = typeof(int);
                        }
                        else // Előjel nélküli
                        {
                            if(bytesTmp == 1) // 8 bites előjel nélküli egész
                                settingsData.setCommands[i].sections[j - sectionCounter].type = typeof(byte);
                            else if(bytesTmp == 2) // 16 bites előjel nélküli egész
                                settingsData.setCommands[i].sections[j - sectionCounter].type = typeof(ushort);
                            else // 32 bites előjel nélküli egész
                                settingsData.setCommands[i].sections[j - sectionCounter].type = typeof(uint);
                        }
                    }
                    
                    for (int k = bitCounter; k < bitCounter + bitsTmp; k++) // Végigmegyünk az adathoz tartozó biteken
                    {
                        nameTmp = Convert.ToString(sBitsTable[k, 0]); // Az adat nevének a beolvasása
                        if (nameTmp.Length < 1 || 32 < nameTmp.Length) // Ha túl rövid, vagy túl hosszú a név (minimum 1, maximum
                        { // 32 karakter),
                            faultText = "The length of the set bits's name in the L" + (k + 2).ToString() + Environment.NewLine +
                                "cell should be minimum 1 character, and maximum 32 characters!";
                            return false; // akkor jelezzük, hogy hibás a táblázat
                        }

                        bool isUsedTmp;
                        try
                        {
                            isUsedTmp = Convert.ToBoolean(sBitsTable[k, 1]); // Annak beolvasása, hogy a bit használatban van-e
                        }
                        catch // Ha a beolvasott adat nem logikai érték,
                        {
                            faultText = "The 'is used' value of the set section in the M" + (k + 2).ToString() + Environment.NewLine +
                                "cell should be true or false!";
                            return false; // akkor jelezzük, hogy hibás a táblázat
                        }

                        settingsData.setCommands[i].sections[j - sectionCounter].bits.Add(new SetBit()); // Létrehozzuk az új bitet,
                        settingsData.setCommands[i].sections[j - sectionCounter].bits[k - bitCounter].name = nameTmp; // majd lementjük
                        settingsData.setCommands[i].sections[j - sectionCounter].bits[k - bitCounter].isUsed = isUsedTmp; // az adatait
                    }
                    
                    bitCounter += bitsTmp; // A bitek számlálóját eltoljuk a már kiolvasott bitek számával
                }

                sectionCounter += sectionsTmp; // Az adatok számlálóját eltoljuk a már kiolvasott adatok számával
            }

            settingsData.logFileHeader = createLogFileHeader(); // Létrehozzuk a logfájl fejlécét

            return true;
        }

        /// <summary>Ez a függvény készíti el a logfájl fejlécét a paramétertáblázat adatai alapján</summary>
        /// <returns>A logfájl fejléce</returns>
        private string createLogFileHeader()
        {
            string header = "";

            header = "System ID: " + settingsData.systemId.ToString() + Environment.NewLine; // Az általános adatok,
            header += "Time;ID"; // amelyek minden logfájlban megtalálhatóak

            for (byte i = 0; i < settingsData.numOfQueryCommands; i++) // Végigmegyünk a lekérdezhető parancsokon,
            {
                for(byte j = 0; j < settingsData.queryCommands[i].numOfSections; j++) // és a parancsokhoz tartozó adatokon
                {
                    if(settingsData.queryCommands[i].sections[j].save) // Ha az adott adatot le kell menteni a log fájlba,
                        header += ";" + settingsData.queryCommands[i].sections[j].name; // akkor a nevét hozzáadjuk a fejléchez
                }
            }
     
            header += Environment.NewLine;

            return header;
        }

        /// <summary>Ez a függvény megnézi, hogy a kapott e-mail cím szintaktikailag helyes-e</summary>
        /// <param name="address">A vizsgálandó e-mail cím</param>
        /// <returns>Igaz, ha szintaktikailag helyes a vizsgálandó e-mail cím</returns>
        /// <remarks>A kód forrása: http://www.codeproject.com/Articles/22777/Email-Address-Validation-Using-Regular-Expression </remarks>
        private bool ValidateMail(string address)
        {
            const string matchEmailPattern = // A minta a sztringhez
                @"^(([\w-]+\.)+[\w-]+|([a-zA-Z]{1}|[\w-]{2,}))@" +
                @"((([0-1]?[0-9]{1,2}|25[0-5]|2[0-4][0-9])\.([0-1]?[0-9]{1,2}|25[0-5]|2[0-4][0-9])\." +
                @"([0-1]?[0-9]{1,2}|25[0-5]|2[0-4][0-9])\.([0-1]?[0-9]{1,2}|25[0-5]|2[0-4][0-9])){1}|" +
                @"([a-zA-Z]+[\w-]+\.)+[a-zA-Z]{2,4})$";

            if (address != null) return Regex.IsMatch(address, matchEmailPattern); // Ellenőrizzük a kapott címet, és ha helyes, igazat,
            else return false; // különben pedig hamisat adunk vissza
        }

    }
}
