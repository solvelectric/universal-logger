using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace universal_logger
{
    /// <summary>Ez az osztály írja le az eszközt, illetve kezeli le a kommunikációt vele</summary>
    /// <remarks>
    /// Az osztály nem rendelkezik teljes kivételkezeléssel, így ezt felsőbb szinten kell megoldani.
    /// Az egyes függvények leírásában található, hogy mely kivételeket kezelik le.
    /// </remarks>
    class Device
    {
        /// <summary>Az aktuális lekérdezés dátuma</summary>
        public DateTime currentDateTime;

        /// <summary>Az eszköz azonosítója</summary>
        public byte id;

        /// <summary>Az eszköz típus- illetve működési információja</summary>
        public byte type;

        /// <summary>Az eszköz firmware verziószáma ééhhnn formátumban (é: év, h: hónap, n: nap)</summary>
        /// <remarks>
        /// Például ha a szám értéke 130204, akkor az azt jelenti, hogy a firmware verziószáma 130204, azaz
        /// a 2013. február 4-ei verzió
        /// </remarks>
        public uint firmwareVersion;

        /// <summary>A lekérdezhető adatokat tartalmazó lista</summary>
        /// <remarks>A külső lista tartalmazza a lekérdező parancsokat, a belső lista pedig az azon belüli adatokat.</remarks>
        public List<List<Object>> queryData;

        /// <summary>A beállítható adatokat tartalmazó lista</summary>
        /// <remarks>A külső lista tartalmazza a beállító parancsokat, a belső lista pedig az azon belüli adatokat.</remarks>
        public List<List<Object>> setData;

        /// <summary>A kommunikáció során maximum ennyiszer próbálkozunk újra, mielőtt hibát jelzünk</summary>
        public const int maxNumOfTimeouts = 3;

        /// <summary>A legutolsó kommunikáció alkalmával kapott válasz szöveges formátumban</summary>
        /// <remarks>A szöveg a kapott bájtokat tartamlazza hexadecimális formában, vesszővel elválasztva.</remarks>
        public string lastReceivedData;

        /// <summary>A legutolsó kommunikáció által kapott válasz dekódolt változata</summary>
        public string lastReceivedDataDecoded;

        /// <summary>Konstruktor az eszköz azonosítójával és a parancsok beállításaival</summary>
        /// <param name="idTmp">Az eszköz azonosítója</param>
        /// <param name="opParamTmp">A használt paramétertábla</param>
        public Device(byte idTmp, OperatingParameters opParamTmp)
        {
            id = idTmp; // Az eszköz azonosítójának a beállítása

            currentDateTime = new DateTime(0); // Az idő inicializálása

            queryData = new List<List<Object>>(opParamTmp.settingsData.numOfQueryCommands); // Létrehozzuk a lekérdező parancsok listáját,
            for (byte i = 0; i < opParamTmp.settingsData.numOfQueryCommands; i++) // majd végigmegyünk rajtuk,
            {
                queryData.Add(new List<Object>(opParamTmp.settingsData.queryCommands[i].numOfSections)); // és létrehozzuk az adatok listáját
                for (byte j = 0; j < opParamTmp.settingsData.queryCommands[i].numOfSections; j++) // Végezetül végigmegyünk az adatokon,
                {
                    Type typeTmp = opParamTmp.settingsData.queryCommands[i].sections[j].type; // és megvizsgáljuk azok típusát
                    if (typeTmp == typeof(DateTime)) // Az adat típusa dátum/idő
                        queryData[i].Add(new DateTime(0));
                    else if (typeTmp == typeof(sbyte)) // Az adat 8 bites előjeles egész
                        queryData[i].Add(new sbyte());
                    else if (typeTmp == typeof(short)) // Az adat 16 bites előjeles egész
                        queryData[i].Add(new short());
                    else if (typeTmp == typeof(int)) // Az adat 32 bites előjeles egész
                        queryData[i].Add(new int());
                    else if (typeTmp == typeof(byte)) // Az adat 8 bites előjel nélküli egész
                        queryData[i].Add(new byte());
                    else if (typeTmp == typeof(ushort)) // Az adat 16 bites előjel nélküli egész
                        queryData[i].Add(new ushort());
                    else queryData[i].Add(new uint()); // Az adat 32 bites előjel nélküli egész
                }
            }

            setData = new List<List<Object>>(opParamTmp.settingsData.numOfSetCommands); // Létrehozzuk a beállító parancsok listáját,
            for (byte i = 0; i < opParamTmp.settingsData.numOfSetCommands; i++) // majd végigmegyünk rajtuk,
            {
                setData.Add(new List<Object>(opParamTmp.settingsData.setCommands[i].numOfSections)); // és létrehozzuk az adatok listáját
                for (byte j = 0; j < opParamTmp.settingsData.setCommands[i].numOfSections; j++) // Végezetül végigmegyünk az adatokon,
                {
                    Type typeTmp = opParamTmp.settingsData.setCommands[i].sections[j].type; // és megvizsgáljuk azok típusát
                    if (typeTmp == typeof(DateTime)) // Az adat típusa dátum/idő
                        setData[i].Add(new DateTime(0));
                    else if (typeTmp == typeof(sbyte)) // Az adat 8 bites előjeles egész
                        setData[i].Add(new sbyte());
                    else if (typeTmp == typeof(short)) // Az adat 16 bites előjeles egész
                        setData[i].Add(new short());
                    else if (typeTmp == typeof(int)) // Az adat 32 bites előjeles egész
                        setData[i].Add(new int());
                    else if (typeTmp == typeof(byte)) // Az adat 8 bites előjel nélküli egész
                        setData[i].Add(new byte());
                    else if (typeTmp == typeof(ushort)) // Az adat 16 bites előjel nélküli egész
                        setData[i].Add(new ushort());
                    else setData[i].Add(new uint()); // Az adat 32 bites előjel nélküli egész
                }
            }

            return;
        }

        /// <summary>A kapott tömb szöveggé konvertálása</summary>
        /// <param name="data">A bájtokat tartalmazó tömb</param>
        /// <returns>Az átkonvertált szöveg</returns>
        /// <remarks>A szöveg a bájtokat hexadecimális formában tartalmazza, vesszővel elválasztva</remarks>
        private string ConvertByteArrayToString(byte[] data)
        {
            string str = ""; // A szövegdoboz kiürítése

            for (ushort i = 0; i < data.GetLength(0); i++) // Végigmegyünk az összes bájton,
            {
                str += Convert.ToString(data[i], 16); // és betesszük őket a szövegbe
                if (i != (data.GetLength(0) - 1)) str += ", ";
            }

            return str;
        }

        /// <summary>Az eszköz hívása</summary>
        /// <param name="port">A használandó COM port</param>
        /// <param name="pid">A használandó [PID] bájt (csomagazonosító)</param>
        /// <returns>Igaz, ha a kommunikáció sikeres volt, az egység elérhető</returns>
        /// <remarks>
        /// Küldött csomag: [START][PID][ID][HEADER][CHK][STOP], fogadott csomag: [START][PID][ID][HEADER][CHK][STOP]
        /// Ez a függvény csak az időtúllépés kivételét kezeli, a többi kivételt azonban továbbdobja, ezért a kivételkezelést felsőbb
        /// szinten kell megoldani.
        /// </remarks>
        public bool Calling(Communication port, byte pid)
        {
            byte[] receive = new byte[6]; // Ide mentjük a kapott bájtokat
            byte[] send = new byte[6]; // Ide másoljuk a küldendő bájtokat

            send[0] = Constants.masterToSlaveStart; // [START]
            send[1] = pid; // [PID]
            send[2] = id; // [ID]
            send[3] = Constants.headerCalling; // [HEADER]
            send[4] = (byte)(0xFF - (send[2] + send[3]) + 0x01); // [CHK] = 0xFF - ([ID] + [HEADER] + [DATA1] + ... + [DATAn]) + 0x01
            send[5] = Constants.masterToSlaveStop; // [STOP]           

            bool tryAgain = true;
            int numOfTries = 0;
            while (tryAgain)
            {
                try
                {
                    lastReceivedData = ""; // Töröljük az utolsó kommunikáció válaszát
                    receive = port.SendPacket(send, 6); // A csomag elküldése, amelyre nullabájtos választ várunk
                    lastReceivedData = ConvertByteArrayToString(receive); // A kapott választ átkonvertáljuk szöveggé
                    tryAgain = false; // Ha az eszköz válaszolt, akkor nem próbálkozunk tovább
                }
                catch (TimeoutException) // Ha időtúllépés történt,
                {
                    numOfTries++; // akkor növeljük a próbálkozások számát
                    if (maxNumOfTimeouts <= numOfTries) // Ha elértük a maximális próbálkozások számát,
                    {
                        lastReceivedDataDecoded = "The device has not answered!";
                        return false; // akkor jelezzük ezt
                    }
                }
                catch // Ha bármilyen más kivétel keletkezett,
                {
                    throw; // akkor azt továbbdobjuk
                }
            }

            bool packetOk;
            if (send[2] == 0) // Ha a csomagot az általános címre küldtük,
                packetOk = port.IsPacketOk(receive, send[1], send[3]); // akkor nem vizsgáljuk az eszközazonosítót,
            else // ellenkező esetben azonban megköveteljük,
                packetOk = port.IsPacketOk(receive, send[1], send[2], send[3]); // hogy a csomag a helyes címről érkezzen

            if (packetOk) // A csomag rendben van
            {
                lastReceivedDataDecoded = "The device has answered!";
                if (send[2] == 0) // Ha a csomagot az általános címre küldtük, akkor kiírjuk az eszközazonosítót is
                    lastReceivedDataDecoded += " The address of the device is " + receive[2].ToString() + ".";
                return true; // A kommunikáció sikeres volt, az eszköz elérhető
            }
            else // A kommunikáció sikertelen volt
            {
                lastReceivedDataDecoded = "The incoming packet was incorrect!";
                return false;
            }
        }

        /// <summary>Az eszköz újraindítása</summary>
        /// <param name="port">A használandó COM port</param>
        /// <param name="pid">A használandó [PID] bájt (csomagazonosító)</param>
        /// <returns>Igaz, ha a kommunikáció sikeres volt, az egység elérhető</returns>
        /// <remarks>
        /// Küldött csomag: [START][PID][ID][HEADER][CHK][STOP], fogadott csomag: [START][PID][ID][HEADER][DATA1][CHK][STOP]
        /// Ez a függvény csak az időtúllépés kivételét kezeli, a többi kivételt azonban továbbdobja, ezért a kivételkezelést felsőbb
        /// szinten kell megoldani.
        /// </remarks>
        public bool Reset(Communication port, byte pid)
        {
            byte[] receive = new byte[7]; // Ide mentjük a kapott bájtokat
            byte[] send = new byte[6]; // Ide másoljuk a küldendő bájtokat

            send[0] = Constants.masterToSlaveStart; // [START]
            send[1] = pid; // [PID]
            send[2] = id; // [ID]
            send[3] = Constants.headerReset; // [HEADER]
            send[4] = (byte)(0xFF - (send[2] + send[3]) + 0x01); // [CHK] = 0xFF - ([ID] + [HEADER] + [DATA1] + ... + [DATAn]) + 0x01
            send[5] = Constants.masterToSlaveStop; // [STOP]

            bool tryAgain = true;
            int numOfTries = 0;
            while (tryAgain)
            {
                try
                {
                    if (send[2] != 0) // Ha a parancsot nem az általános címre küldtük, akkor választ várunk az eszköztől
                    {
                        lastReceivedData = ""; // Töröljük az utolsó kommunikáció válaszát
                        receive = port.SendPacket(send, 7); // A csomag elküldése, amelyre egybájtos választ várunk
                        lastReceivedData = ConvertByteArrayToString(receive); // A kapott választ átkonvertáljuk szöveggé
                        tryAgain = false; // Ha az eszköz válaszolt, akkor nem próbálkozunk tovább
                    }
                    else // Ha a parancsot az általános címre küldtük, akkor az eszköz nem küld vissza semmit sem,
                    {
                        port.SendPacket(send); // így a csomag elküldése után mi sem várunk választ
                        lastReceivedData = "The device will not answer, so we assume that the command was successful.";
                        lastReceivedData = ""; // Beállítjuk a standard válaszokat,
                        return true; // majd jelezzük, hogy a parancs sikeresen végre lett hajtva
                    }
                }
                catch (TimeoutException) // Ha időtúllépés történt,
                {
                    numOfTries++; // akkor növeljük a próbálkozások számát
                    if (maxNumOfTimeouts <= numOfTries) // Ha elértük a maximális próbálkozások számát,
                    {
                        lastReceivedDataDecoded = "The device has not answered!";
                        return false; // akkor jelezzük ezt
                    }
                }
                catch // Ha bármilyen más kivétel keletkezett,
                {
                    throw; // akkor azt továbbdobjuk
                }
            }

            if (port.IsPacketOk(receive, send[1], send[2], send[3])) // A csomag rendben van,
            {
                lastReceivedDataDecoded = port.DecodeAnswerData(receive[4]); // A válasz dekódolása
                return port.DecodeAnswerSuccess(receive[4]);
            }
            else // A kommunikáció sikertelen volt
            {
                lastReceivedDataDecoded = "The incoming packet was incorrect!";
                return false;
            }
        }

        /// <summary>Az eszköz típus- illetve működési információjának a lekérdezése</summary>
        /// <param name="port">A használandó COM port</param>
        /// <param name="pid">A használandó [PID] bájt (csomagazonosító)</param>
        /// <returns>Igaz, ha a kommunikáció sikeres volt, az egység elérhető</returns>
        /// <remarks>
        /// Küldött csomag: [START][PID][ID][HEADER][CHK][STOP], fogadott csomag: [START][PID][ID][HEADER][DATA1][CHK][STOP]
        /// Ez a függvény csak az időtúllépés kivételét kezeli, a többi kivételt azonban továbbdobja, ezért a kivételkezelést felsőbb
        /// szinten kell megoldani.
        /// </remarks>
        public bool AskType(Communication port, byte pid)
        {
            byte[] receive = new byte[7]; // Ide mentjük a kapott bájtokat
            byte[] send = new byte[6]; // Ide másoljuk a küldendő bájtokat

            send[0] = Constants.masterToSlaveStart; // [START]
            send[1] = pid; // [PID]
            send[2] = id; // [ID]
            send[3] = Constants.headerAskType; // [HEADER]
            send[4] = (byte)(0xFF - (send[2] + send[3]) + 0x01); // [CHK] = 0xFF - ([ID] + [HEADER] + [DATA1] + ... + [DATAn]) + 0x01
            send[5] = Constants.masterToSlaveStop; // [STOP]

            bool tryAgain = true;
            int numOfTries = 0;
            while (tryAgain)
            {
                try
                {
                    if (send[2] != 0) // Ha a parancsot nem az általános címre küldtük, akkor választ várunk az eszköztől
                    {
                        lastReceivedData = ""; // Töröljük az utolsó kommunikáció válaszát
                        receive = port.SendPacket(send, 7); // A csomag elküldése, amelyre egybájtos választ várunk
                        lastReceivedData = ConvertByteArrayToString(receive); // A kapott választ átkonvertáljuk szöveggé
                        tryAgain = false; // Ha az eszköz válaszolt, akkor nem próbálkozunk tovább
                    }
                    else // Ha a parancsot az általános címre küldtük, akkor az eszköz nem küld vissza semmit sem,
                    {
                        port.SendPacket(send); // így a csomag elküldése után mi sem várunk választ
                        lastReceivedData = "The device will not answer, so we assume that the command was successful.";
                        lastReceivedData = ""; // Beállítjuk a standard válaszokat,
                        return true; // majd jelezzük, hogy a parancs sikeresen végre lett hajtva
                    }
                }
                catch (TimeoutException) // Ha időtúllépés történt,
                {
                    numOfTries++; // akkor növeljük a próbálkozások számát
                    if (maxNumOfTimeouts <= numOfTries) // Ha elértük a maximális próbálkozások számát,
                    {
                        lastReceivedDataDecoded = "The device has not answered!";
                        return false; // akkor jelezzük ezt
                    }
                }
                catch // Ha bármilyen más kivétel keletkezett,
                {
                    throw; // akkor azt továbbdobjuk
                }
            }

            if (port.IsPacketOk(receive, send[1], send[2], send[3])) // A csomag rendben van,
            {
                lastReceivedDataDecoded = "The packet was received successfully!";
                type = receive[4]; // Az eszköz típus- illetve működési információjának lementése
                return true;
            }
            else // A kommunikáció sikertelen volt
            {
                lastReceivedDataDecoded = "The incoming packet was incorrect!";
                return false;
            }
        }

        /// <summary>Az eszköz firmware verziószámának a lekérdezése</summary>
        /// <param name="port">A használandó COM port</param>
        /// <param name="pid">A használandó [PID] bájt (csomagazonosító)</param>
        /// <returns>Igaz, ha a kommunikáció sikeres volt</returns>
        /// <remarks>
        /// Küldött csomag: [START][PID][ID][HEADER][CHK][STOP],
        /// fogadott csomag: [START][PID][ID][HEADER][DATA1][DATA2][DATA3][DATA4][CHK][STOP]
        /// Ez a függvény csak az időtúllépés kivételét kezeli, a többi kivételt azonban továbbdobja, ezért a kivételkezelést felsőbb
        /// szinten kell megoldani.
        /// </remarks>
        public bool AskFirmwareVersion(Communication port, byte pid)
        {
            byte[] receive = new byte[10]; // Ide mentjük a kapott bájtokat
            byte[] send = new byte[6]; // Ide másoljuk a küldendő bájtokat

            send[0] = Constants.masterToSlaveStart; // [START]
            send[1] = pid; // [PID]
            send[2] = id; // [ID]
            send[3] = Constants.headerAskFwVersion; // [HEADER]
            send[4] = (byte)(0xFF - (send[2] + send[3]) + 0x01); // [CHK] = 0xFF - ([ID] + [HEADER] + [DATA1] + ... + [DATAn]) + 0x01
            send[5] = Constants.masterToSlaveStop; // [STOP]

            bool tryAgain = true;
            int numOfTries = 0;
            while (tryAgain)
            {
                try
                {
                    if (send[2] != 0) // Ha a parancsot nem az általános címre küldtük, akkor választ várunk az eszköztől
                    {
                        lastReceivedData = ""; // Töröljük az utolsó kommunikáció válaszát
                        receive = port.SendPacket(send, 10); // A csomag elküldése, amelyre négybájtos választ várunk
                        lastReceivedData = ConvertByteArrayToString(receive); // A kapott választ átkonvertáljuk szöveggé
                        tryAgain = false; // Ha az eszköz válaszolt, akkor nem próbálkozunk tovább
                    }
                    else // Ha a parancsot az általános címre küldtük, akkor az eszköz nem küld vissza semmit sem,
                    {
                        port.SendPacket(send); // így a csomag elküldése után mi sem várunk választ
                        lastReceivedData = "The device will not answer, so we assume that the command was successful.";
                        lastReceivedData = ""; // Beállítjuk a standard válaszokat,
                        return true; // majd jelezzük, hogy a parancs sikeresen végre lett hajtva
                    }
                }
                catch (TimeoutException) // Ha időtúllépés történt,
                {
                    numOfTries++; // akkor növeljük a próbálkozások számát
                    if (maxNumOfTimeouts <= numOfTries) // Ha elértük a maximális próbálkozások számát,
                    {
                        lastReceivedDataDecoded = "The device has not answered!";
                        return false; // akkor jelezzük ezt
                    }
                }
                catch // Ha bármilyen más kivétel keletkezett,
                {
                    throw; // akkor azt továbbdobjuk
                }
            }

            if (port.IsPacketOk(receive, send[1], send[2], send[3])) // A csomag rendben van, ezért lemenjük a kapott verziószámot
            {
                lastReceivedDataDecoded = "The packet was received successfully!";

                firmwareVersion = (uint)((((uint)receive[4]) << 24) | (((uint)receive[5]) << 16) | // A verziószám
                    (((uint)receive[6]) << 8) | (uint)receive[7]);
                return true; // A kommunikáció sikeres volt
            }
            else // A kommunikáció sikertelen volt
            {
                lastReceivedDataDecoded = "The incoming packet was incorrect!";
                return false;
            }
        }

        /// <summary> Az eszköz címének a beállítása</summary>
        /// <param name="port">A használandó COM port</param>
        /// <param name="pid">A használandó [PID] bájt (csomagazonosító)</param>
        /// <param name="data1">Az egység új címe</param>
        /// <returns>Igaz, ha az érték sikeresen be lett állítva (vagy már be volt állítva)</returns>
        /// <remarks>
        /// Küldött csomag: [START][PID][ID][HEADER][DATA1][CHK][STOP], fogadott csomag: [START][PID][ID][HEADER][DATA1][CHK][STOP]
        /// Ez a függvény csak az időtúllépés kivételét kezeli, a többi kivételt azonban továbbdobja, ezért a kivételkezelést felsőbb
        /// szinten kell megoldani.
        /// </remarks>
        public bool SendNewId(Communication port, byte pid, byte data1)
        {
            byte[] receive = new byte[7]; // Ide mentjük a kapott bájtokat
            byte[] send = new byte[7]; // Ide másoljuk a küldendő bájtokat

            send[0] = Constants.masterToSlaveStart; // [START]
            send[1] = pid; // [PID]
            send[2] = id; // [ID]
            send[3] = Constants.headerSetAddress; // [HEADER]
            send[4] = data1; // [DATA1]
            // [CHK] = 0xFF - ([ID] + [HEADER] + [DATA1] + ... + [DATAn]) + 0x01
            send[5] = (byte)(0xFF - (send[2] + send[3] + send[4]) + 0x01);
            send[6] = Constants.masterToSlaveStop; // [STOP]

            bool tryAgain = true;
            int numOfTries = 0;
            while (tryAgain)
            {
                try
                {
                    if (send[2] != 0) // Ha a parancsot nem az általános címre küldtük, akkor választ várunk az eszköztől
                    {
                        lastReceivedData = ""; // Töröljük az utolsó kommunikáció válaszát
                        receive = port.SendPacket(send, 7); // A csomag elküldése, amelyre egybájtos választ várunk
                        lastReceivedData = ConvertByteArrayToString(receive); // A kapott választ átkonvertáljuk szöveggé
                        tryAgain = false; // Ha az eszköz válaszolt, akkor nem próbálkozunk tovább
                    }
                    else // Ha a parancsot az általános címre küldtük, akkor az eszköz nem küld vissza semmit sem,
                    {
                        port.SendPacket(send); // így a csomag elküldése után mi sem várunk választ
                        lastReceivedData = "The device will not answer, so we assume that the command was successful.";
                        lastReceivedData = ""; // Beállítjuk a standard válaszokat,
                        id = send[4]; // lementjük az új értéket,
                        return true; // majd jelezzük, hogy a parancs sikeresen végre lett hajtva
                    }
                }
                catch (TimeoutException) // Ha időtúllépés történt,
                {
                    numOfTries++; // akkor növeljük a próbálkozások számát
                    if (maxNumOfTimeouts <= numOfTries) // Ha elértük a maximális próbálkozások számát,
                    {
                        lastReceivedDataDecoded = "The device has not answered!";
                        return false; // akkor jelezzük ezt
                    }
                }
                catch // Ha bármilyen más kivétel keletkezett,
                {
                    throw; // akkor azt továbbdobjuk
                }
            }

            if (port.IsPacketOk(receive, send[1], send[2], send[3])) // A csomag rendben van
            {
                lastReceivedDataDecoded = port.DecodeAnswerData(receive[4]); // A válasz dekódolása
                if (port.DecodeAnswerSuccess(receive[4])) // Ha a kommunikáció sikeres volt,
                {
                    id = send[4]; // akkor lementjük az új értéket
                    return true;
                }
                else return false; // A cím átállítása sikertelen volt
            }
            else // A kommunikáció sikertelen volt
            {
                lastReceivedDataDecoded = "The incoming packet was incorrect!";
                return false;
            }
        }

        /// <summary>Általános lekérdezés az eszköz felé</summary>
        /// <param name="port">A használandó COM port</param>
        /// <param name="pid">A használandó [PID] bájt (csomagazonosító)</param>
        /// <param name="commandNumber">A parancs sorszáma</param>
        /// <param name="opParamTmp">A használt paramétertáblázat</param>
        /// <returns>Igaz, ha a lekérdezés sikeres volt</returns>
        /// <remarks>
        /// Küldött csomag: [START][PID][ID][HEADER][CHK][STOP], fogadott csomag: [START][PID][ID][HEADER][DATA1]...[DATAn][CHK][STOP]
        /// Ez a függvény csak az időtúllépés kivételét kezeli, a többi kivételt azonban továbbdobja, ezért a kivételkezelést felsőbb
        /// szinten kell megoldani.
        /// </remarks>
        public bool AskGeneral(Communication port, byte pid, byte commandNumber, OperatingParameters opParamTmp)
        {
            ushort receiveSize = (ushort)(6 + opParamTmp.settingsData.queryCommands[commandNumber].dataBytes); // A válaszcsomag mérete

            byte[] receive = new byte[receiveSize]; // Ide mentjük a kapott bájtokat
            byte[] send = new byte[6]; // Ide másoljuk a küldendő bájtokat

            send[0] = Constants.masterToSlaveStart; // [START]
            send[1] = pid; // [PID]
            send[2] = id; // [ID]
            send[3] = opParamTmp.settingsData.queryCommands[commandNumber].code; // [HEADER]
            send[4] = (byte)(0xFF - (send[2] + send[3]) + 0x01); // [CHK] = 0xFF - ([ID] + [HEADER] + [DATA1] + ... + [DATAn]) + 0x01
            send[5] = Constants.masterToSlaveStop; // [STOP]

            bool tryAgain = true;
            int numOfTries = 0;
            while (tryAgain)
            {
                try
                {
                    if (send[2] != 0) // Ha a parancsot nem az általános címre küldtük, akkor választ várunk az eszköztől
                    {
                        lastReceivedData = ""; // Töröljük az utolsó kommunikáció válaszát
                        receive = port.SendPacket(send, receiveSize); // A csomag elküldése, amelyre receiveSize bájtos választ várunk
                        lastReceivedData = ConvertByteArrayToString(receive); // A kapott választ átkonvertáljuk szöveggé
                        tryAgain = false; // Ha az eszköz válaszolt, akkor nem próbálkozunk tovább
                    }
                    else // Ha a parancsot az általános címre küldtük, akkor az eszköz nem küld vissza semmit sem,
                    {
                        port.SendPacket(send); // így a csomag elküldése után mi sem várunk választ
                        lastReceivedData = "The device will not answer, so we assume that the command was successful.";
                        lastReceivedData = ""; // Beállítjuk a standard válaszokat,
                        return true; // majd jelezzük, hogy a parancs sikeresen végre lett hajtva
                    }
                }
                catch (TimeoutException) // Ha időtúllépés történt,
                {
                    numOfTries++; // akkor növeljük a próbálkozások számát
                    if (maxNumOfTimeouts <= numOfTries) // Ha elértük a maximális próbálkozások számát,
                    {
                        lastReceivedDataDecoded = "The device has not answered!";
                        return false; // akkor jelezzük ezt
                    }
                }
                catch // Ha bármilyen más kivétel keletkezett,
                {
                    throw; // akkor azt továbbdobjuk
                }
            }

            if (port.IsPacketOk(receive, send[1], send[2], send[3])) // A csomag rendben van,
            {
                lastReceivedDataDecoded = "The packet was received successfully!";

                ushort j = 4; // Az aktuálisan vizsgált bájt sorszáma
                for (byte i = 0; i < opParamTmp.settingsData.queryCommands[commandNumber].numOfSections; i++) // Végigmegyünk az adatokon,
                {
                    Type typeTmp = opParamTmp.settingsData.queryCommands[commandNumber].sections[i].type; // és megvizsgáljuk azok típusát
                    if (typeTmp == typeof(DateTime)) // Az adat típusa dátum/idő
                    {
                        // A dátum/idő BCD formátumban van, ezért át kell alakítani
                        ushort tmp = (ushort)((((ushort)receive[j]) << 8) | (ushort)receive[j + 1]); // A dátumból az év
                        j += 2; // Ugrunk a következő vizsgálandó bájtra
                        ushort year = (ushort)(tmp & 0x000f); // Az egyesek
                        tmp >>= 4; // Az egyesek törlése
                        year += (ushort)((ushort)(tmp & 0x000f) * 10); // A tizesek
                        tmp >>= 4; // A tizesek törlése
                        year += (ushort)((ushort)(tmp & 0x000f) * 100); // A százasok
                        tmp >>= 4; // A százasok törlése
                        year += (ushort)(tmp * 1000); // Az ezresek

                        byte month = (byte)(((byte)(((byte)(receive[j] >> 4)) * 10)) + ((byte)(receive[j] & 0x0f))); // A dátumból a hónap
                        j++; // Ugrunk a következő vizsgálandó bájtra
                        byte day = (byte)(((byte)(((byte)(receive[j] >> 4)) * 10)) + ((byte)(receive[j] & 0x0f))); // A dátumból a nap
                        j++; // Ugrunk a következő vizsgálandó bájtra
                        byte hour = (byte)(((byte)(((byte)(receive[j] >> 4)) * 10)) + ((byte)(receive[j] & 0x0f))); // Az időből az óra
                        j++; // Ugrunk a következő vizsgálandó bájtra
                        byte min = (byte)(((byte)(((byte)(receive[j] >> 4)) * 10)) + ((byte)(receive[j] & 0x0f))); // Az időből a perc
                        j++; // Ugrunk a következő vizsgálandó bájtra
                        byte sec = (byte)(((byte)(((byte)(receive[j] >> 4)) * 10)) + ((byte)(receive[j] & 0x0f))); // Az időből a másodperc
                        j++; // Ugrunk a következő vizsgálandó bájtra

                        queryData[commandNumber][i] = new DateTime(year, month, day, hour, min, sec); // Lementjük az értéket

                        if (opParamTmp.settingsData.queryCommands[commandNumber].sections[i].save) // Ha az adatot le kell menteni logfájlba,
                            currentDateTime = new DateTime(year, month, day, hour, min, sec); // akkor frissítjük a lekérdezés dátumát is
                    }
                    else if (typeTmp == typeof(sbyte)) // Az adat 8 bites előjeles egész
                    {
                        queryData[commandNumber][i] = (sbyte)(receive[j]); // Lementjük az értéket,
                        j++; // és ugrunk a következő vizsgálandó bájtra
                    }
                    else if (typeTmp == typeof(short)) // Az adat 16 bites előjeles egész
                    {
                        queryData[commandNumber][i] = (short)((((ushort)receive[j]) << 8) | (ushort)receive[j + 1]); // Lementjük az értéket,
                        j += 2; // és ugrunk a következő vizsgálandó bájtra
                    }
                    else if (typeTmp == typeof(int)) // Az adat 32 bites előjeles egész
                    {
                        queryData[commandNumber][i] = (int)((((uint)receive[j]) << 24) | (((uint)receive[j + 1]) << 16) |
                            (((uint)receive[j + 2]) << 8) | (uint)receive[j + 3]); // Lementjük az értéket,
                        j += 4; // és ugrunk a következő vizsgálandó bájtra
                    }
                    else if (typeTmp == typeof(byte)) // Az adat 8 bites előjel nélküli egész
                    {
                        queryData[commandNumber][i] = receive[j]; // Lementjük az értéket,
                        j++; // és ugrunk a következő vizsgálandó bájtra
                    }
                    else if (typeTmp == typeof(ushort)) // Az adat 16 bites előjel nélküli egész
                    {
                        queryData[commandNumber][i] = (ushort)((((ushort)receive[j]) << 8) | (ushort)receive[j + 1]); // Lementjük
                        j += 2; // az értéket, és ugrunk a következő vizsgálandó bájtra
                    }
                    else // Az adat 32 bites előjel nélküli egész
                    {
                        queryData[commandNumber][i] = (uint)((((uint)receive[j]) << 24) | (((uint)receive[j + 1]) << 16) |
                            (((uint)receive[j + 2]) << 8) | (uint)receive[j + 3]); // Lementjük az értéket,
                        j += 4; // és ugrunk a következő vizsgálandó bájtra
                    }                    
                }

                return true;
            }
            else // A kommunikáció sikertelen volt
            {
                lastReceivedDataDecoded = "The incoming packet was incorrect!";
                return false;
            }
        }

        /// <summary>Általános parancs küldése az eszköz felé</summary>
        /// <param name="port">A használandó COM port</param>
        /// <param name="pid">A használandó [PID] bájt (csomagazonosító)</param>
        /// <param name="commandNumber">A parancs sorszáma</param>
        /// <param name="opParamTmp">A használt paramétertáblázat</param>
        /// <param name="setDataTmp">Az elküldendő adatokat tartalmazó objektum</param>
        /// <returns>Igaz, ha az adatok beállítása sikeres volt</returns>
        /// <remarks>
        /// Küldött csomag: [START][PID][ID][HEADER][DATA1]...[DATAn][CHK][STOP],
        /// fogadott csomag: [START][PID][ID][HEADER][DATA1][CHK][STOP], de csak abban az esetben, ha standard válasz érkezik a csomagra.
        /// Ez a függvény csak az időtúllépés kivételét kezeli, a többi kivételt azonban továbbdobja, ezért a kivételkezelést felsőbb
        /// szinten kell megoldani.
        /// </remarks>
        public bool SendGeneral(Communication port, byte pid, byte commandNumber, OperatingParameters opParamTmp, SetData setDataTmp)
        {
            ushort sendSize = (ushort)(6 + opParamTmp.settingsData.setCommands[commandNumber].dataBytes); // A küldött csomag mérete

            byte[] receive = new byte[7]; // Ide mentjük a kapott bájtokat
            byte[] send = new byte[sendSize]; // Ide másoljuk a küldendő bájtokat

            send[0] = Constants.masterToSlaveStart; // [START]
            send[1] = pid; // [PID]
            send[2] = id; // [ID]
            send[3] = opParamTmp.settingsData.setCommands[commandNumber].code; // [HEADER]

            byte chk = (byte)(send[2] + send[3]); // Az ellenőrző összeg inicializálása
            ushort actualIndex = 4; // Az aktuális elem indexe a küldendő bájtok tömbjében
            for (byte i = 0; i < opParamTmp.settingsData.setCommands[commandNumber].numOfSections; i++) // Végigmegyünk az adatokon,
            {
                Type typeTmp = opParamTmp.settingsData.setCommands[commandNumber].sections[i].type; // majd megvizsgáljuk a típusukat
                if (typeTmp == typeof(DateTime)) // Az adat típusa dátum/idő
                {
                    DateTime data = Convert.ToDateTime(setDataTmp.setSections[i].setValue); // Az adat kiolvasása

                    // Mivel a dátumot/időt BCD formátumban kell elküldeni, ezért az adatot át kell alakítani
                    byte data1 = (byte)(data.Year / 1000); // Az évből az ezresek száma
                    data1 <<= 4;
                    data1 |= (byte)((data.Year % 1000) / 100); // Az évből a százasok száma
                    byte data2 = (byte)((data.Year % 100) / 10); // Az évből a tízesek száma
                    data2 <<= 4;
                    data2 |= (byte)(data.Year % 10); // Az évből az egyesek száma

                    byte data3 = (byte)(data.Month / 10); // A hónapból a tízesek száma
                    data3 <<= 4;
                    data3 |= (byte)(data.Month % 10); // A hónapból az egyesek száma

                    byte data4 = (byte)(data.Day / 10); // A napból a tízesek száma
                    data4 <<= 4;
                    data4 |= (byte)(data.Day % 10); // A napból az egyesek száma

                    byte data5 = (byte)(data.Hour / 10); // Az órából a tízesek száma
                    data5 <<= 4;
                    data5 |= (byte)(data.Hour % 10); // Az órából az egyesek száma

                    byte data6 = (byte)(data.Minute / 10); // A percekből a tízesek száma
                    data6 <<= 4;
                    data6 |= (byte)(data.Minute % 10); // A percekből az egyesek száma

                    byte data7 = (byte)(data.Second / 10); // A másodpercekből a tízesek száma
                    data7 <<= 4;
                    data7 |= (byte)(data.Second % 10); // A másodpercekből az egyesek száma

                    send[actualIndex] = data1; // A bájtok lementése,
                    actualIndex++; // és az index léptetése
                    send[actualIndex] = data2;
                    actualIndex++;
                    send[actualIndex] = data3;
                    actualIndex++;
                    send[actualIndex] = data4;
                    actualIndex++;
                    send[actualIndex] = data5;
                    actualIndex++;
                    send[actualIndex] = data6;
                    actualIndex++;
                    send[actualIndex] = data7;
                    actualIndex++;

                    chk += (byte)(data1 + data2 + data3 + data4 + data5 + data6 + data7); // Az ellenőrző összeg frissítése
                }
                else if (typeTmp == typeof(sbyte)) // Az adat 8 bites előjeles egész
                {
                    sbyte data = Convert.ToSByte(setDataTmp.setSections[i].setValue); // Az adat kiolvasása

                    send[actualIndex] = (byte)data; // A bájt lementése,
                    actualIndex++; // és az index léptetése

                    chk += (byte)data; // Az ellenőrző összeg frissítése
                }
                else if (typeTmp == typeof(short)) // Az adat 16 bites előjeles egész
                {
                    short data = Convert.ToInt16(setDataTmp.setSections[i].setValue); // Az adat kiolvasása

                    byte data1 = (byte)(data >> 8); // Az egyes bájtok
                    byte data2 = (byte)data;

                    send[actualIndex] = data1; // A bájtok lementése,
                    actualIndex++; // és az index léptetése
                    send[actualIndex] = data2;
                    actualIndex++;

                    chk += (byte)(data1 + data2); // Az ellenőrző összeg frissítése
                }
                else if (typeTmp == typeof(int)) // Az adat 32 bites előjeles egész
                {
                    int data = Convert.ToInt32(setDataTmp.setSections[i].setValue); // Az adat kiolvasása

                    byte data1 = (byte)(data >> 24); // Az egyes bájtok
                    byte data2 = (byte)(data >> 16);
                    byte data3 = (byte)(data >> 8);
                    byte data4 = (byte)data;

                    send[actualIndex] = data1; // A bájtok lementése,
                    actualIndex++; // és az index léptetése
                    send[actualIndex] = data2;
                    actualIndex++;
                    send[actualIndex] = data3;
                    actualIndex++;
                    send[actualIndex] = data4;
                    actualIndex++;

                    chk += (byte)(data1 + data2 + data3 + data4); // Az ellenőrző összeg frissítése
                }
                else if (typeTmp == typeof(byte)) // Az adat 8 bites előjel nélküli egész
                {
                    byte data = Convert.ToByte(setDataTmp.setSections[i].setValue); // Az adat kiolvasása

                    send[actualIndex] = data; // A bájt lementése,
                    actualIndex++; // és az index léptetése

                    chk += data; // Az ellenőrző összeg frissítése
                }
                else if (typeTmp == typeof(ushort)) // Az adat 16 bites előjel nélküli egész
                {
                    ushort data = Convert.ToUInt16(setDataTmp.setSections[i].setValue); // Az adat kiolvasása

                    byte data1 = (byte)(data >> 8); // Az egyes bájtok
                    byte data2 = (byte)data;

                    send[actualIndex] = data1; // A bájtok lementése,
                    actualIndex++; // és az index léptetése
                    send[actualIndex] = data2;
                    actualIndex++;

                    chk += (byte)(data1 + data2); // Az ellenőrző összeg frissítése
                }
                else // Az adat 32 bites előjel nélküli egész
                {
                    uint data = Convert.ToUInt32(setDataTmp.setSections[i].setValue); // Az adat kiolvasása

                    byte data1 = (byte)(data >> 24); // Az egyes bájtok
                    byte data2 = (byte)(data >> 16);
                    byte data3 = (byte)(data >> 8);
                    byte data4 = (byte)data;

                    send[actualIndex] = data1; // A bájtok lementése,
                    actualIndex++; // és az index léptetése
                    send[actualIndex] = data2;
                    actualIndex++;
                    send[actualIndex] = data3;
                    actualIndex++;
                    send[actualIndex] = data4;
                    actualIndex++;

                    chk += (byte)(data1 + data2 + data3 + data4); // Az ellenőrző összeg frissítése
                }
            }

            send[actualIndex] = (byte)(0xFF - (chk) + 0x01); // [CHK] = 0xFF - ([ID] + [HEADER] + [DATA1] + ... + [DATAn]) + 0x01
            actualIndex++;
            send[actualIndex] = Constants.masterToSlaveStop; // [STOP]

            bool tryAgain = true;
            int numOfTries = 0;
            while (tryAgain)
            {
                try
                {
                    if (opParamTmp.settingsData.setCommands[commandNumber].standardAnswer && // Ha az eszköz standard egybájtos választ küld,
                        send[2] != 0) // és a parancsot nem az általános címre küldtük, akkor választ várunk az eszköztől
                    {
                        lastReceivedData = ""; // Töröljük az utolsó kommunikáció válaszát
                        receive = port.SendPacket(send, 7); // A csomag elküldése, amelyre receiveSize bájtos választ várunk
                        lastReceivedData = ConvertByteArrayToString(receive); // A kapott választ átkonvertáljuk szöveggé
                        tryAgain = false; // Ha az eszköz válaszolt, akkor nem próbálkozunk tovább
                    }
                    else // Ha az eszköz nem küld vissza semmit sem,
                    {
                        port.SendPacket(send); // akkor a csomag elküldése után mi sem várunk választ,
                        lastReceivedData = "The device will not answer, so we assume that the command was successful.";
                        lastReceivedData = ""; // így beállítjuk a standard válaszokat,
                        for (byte i = 0; i < opParamTmp.settingsData.setCommands[commandNumber].numOfSections; i++) // végigmegyünk az
                            setData[commandNumber][i] = setDataTmp.setSections[i].setValue; // adatokon, és átmásoljuk őket,
                        return true; // majd jelezzük, hogy a parancs sikeresen végre lett hajtva
                    }
                }
                catch (TimeoutException) // Ha időtúllépés történt,
                {
                    numOfTries++; // akkor növeljük a próbálkozások számát
                    if (maxNumOfTimeouts <= numOfTries) // Ha elértük a maximális próbálkozások számát,
                    {
                        lastReceivedDataDecoded = "The device has not answered!";
                        return false; // akkor jelezzük ezt
                    }
                }
                catch // Ha bármilyen más kivétel keletkezett,
                {
                    throw; // akkor azt továbbdobjuk
                }
            }

            if (port.IsPacketOk(receive, send[1], send[2], send[3])) // A csomag rendben van
            {
                lastReceivedDataDecoded = port.DecodeAnswerData(receive[4]); // A válasz dekódolása
                if (port.DecodeAnswerSuccess(receive[4])) // Ha a kommunikáció sikeres volt,
                {
                    for (byte i = 0; i < opParamTmp.settingsData.setCommands[commandNumber].numOfSections; i++) // Végigmegyünk az adatokon,
                        setData[commandNumber][i] = setDataTmp.setSections[i].setValue; // és átmásoljuk őket

                    return true;
                }
                else return false; // A cím átállítása sikertelen volt
            }
            else // A kommunikáció sikertelen volt
            {
                lastReceivedDataDecoded = "The incoming packet was incorrect!";
                return false;
            }
        }

        /// <summary>Ez a függvény összeállít egy sztringet az aktuális adatokból, ami kiírható a logfájlba</summary>
        /// <param name="opParamTmp">A használt paramétertáblázat</param>
        /// <returns>A fájlba írható sor</returns>
        /// <remarks>A visszaadott sor tartalmazza az újsor karaktert is.</remarks>
        public string CreateLogFileLine(OperatingParameters opParamTmp)
        {
            string line = "";

            line += currentDateTime.ToString("HH\\:mm\\:ss"); // A sor inicializálása a lekérdezés idejével,
            line += ";" + id.ToString(); // valamint az eszközazonosítóval

            for (byte i = 0; i < opParamTmp.settingsData.numOfQueryCommands; i++) // Végigmegyünk a lekérdező parancsokon,
            {
                for (byte j = 0; j < opParamTmp.settingsData.queryCommands[i].numOfSections; j++) // és azok adatain
                {
                    if (opParamTmp.settingsData.queryCommands[i].sections[j].save) // Ha az adatot le kell menteni a log fájlba,
                    { // akkor megvizsgáljuk annak típusát
                        if (opParamTmp.settingsData.queryCommands[i].sections[j].decimals == -2) // Az adat hexadecimális
                        {
                            // Lementjük az adat típusát, amely csak előjel nélküli egész lehet
                            Type typeTmp = opParamTmp.settingsData.queryCommands[i].sections[j].type;
                            if (typeTmp == typeof(byte)) // Az adat 8 bites előjel nélküli egész
                                line += ";" + ((byte)(queryData[i][j])).ToString("X2"); // Hozzáadjuk az adatot a sorhoz
                            else if (typeTmp == typeof(ushort)) // Az adat 16 bites előjel nélküli egész
                                line += ";" + ((ushort)(queryData[i][j])).ToString("X4"); // Hozzáadjuk az adatot a sorhoz
                            else // Az adat 32 bites előjel nélküli egész
                                line += ";" + ((uint)(queryData[i][j])).ToString("X8"); // Hozzáadjuk az adatot a sorhoz
                        }
                        else if (opParamTmp.settingsData.queryCommands[i].sections[j].decimals == -1) // Az adat dátum/idő
                        {
                            // A dátum/idő típusú adatokat nem lehet ilyen módon lementeni
                            // Ha egy dátum/idő típusú adatot menteni kell, akkor az azt jelenti, hogy ezt az értéket kell felhasználni,
                            // mint a lekérdezés dátumát
                            // Amennyiben mégis ténylegesen a log fájlba kell menteni egy dátumot, akkor külön kell kezelni a paraméter-
                            // táblában az évet, hónapot, napot, órát, percet és másodpercet
                        }
                        else // Az adat egy "normál" (nem hexadecimális) szám (egész vagy tört)
                        {
                            float value; // A szám értéke
                            string format = "F"; // valamint formátuma

                            value = Convert.ToSingle(queryData[i][j]); // A szám értéke egészként

                            for (byte count = 0; count < opParamTmp.settingsData.queryCommands[i].sections[j].decimals; count++)
                            { // Ahány tizedesjegyet tartalmaz az érték,
                                value /= 10.0F; // annyiszor osztjuk le tízzel
                            }

                            // Kiegészítjük a formátum sztringet a tizedesjegyek számával
                            format += opParamTmp.settingsData.queryCommands[i].sections[j].decimals.ToString();

                            line += ";" + value.ToString(format); // Hozzáadjuk az adatot a sorhoz
                        }
                    }
                    // Ha az adatot nem kell lementeni a logfájlba, akkor nem csinálunk semmit sem
                }
                // Végeztünk az összes, ehhez a parancshoz tartozó adattal
            }

            line += Environment.NewLine; // A sor vége

            return line;
        }
    }
}
