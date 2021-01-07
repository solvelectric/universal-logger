using System;
using System.Collections.Generic;
using System.IO.Ports; // A SerialPort osztályhoz kell
using System.Linq;
using System.Text;
using System.Windows.Forms; // Az üzenetek kiírásához kell (MessageBox)

namespace universal_logger
{
    /// <summary>Ez az osztály kezeli le a soros portot, illetve bonyolítja le a kommunikációt az eszközzel</summary>
    /// <remarks>Az osztály nem rendelkezik kivételkezeléssel, így ezt felsőbb szinten kell megoldani.</remarks>
    class Communication
    {
        /// <summary>A COM portot kezelő osztály</summary>
        private SerialPort _com;

        /// <summary>Konstruktor a port nevével</summary>
        /// <param name="comPortName">Ez a port neve</param>
        /// <param name="baud">A baud rate</param>
        /// <param name="timeout">Az időtúllépés olvasásnál és írásnál</param>
        /// <remarks>A konstruktor nem rendelkezik kivételkezeléssel, így a hívó függvénynek kell ezt megoldania.</remarks>
        public Communication(string comPortName, int baud, int timeout)
        {
            _com = new SerialPort(); // A COM port létrehozása

            // A port paramétereinek beállítása
            _com.PortName = comPortName; // A port neve
            _com.BaudRate = baud; // Baud rate
            _com.Parity = (Parity)Enum.Parse(typeof(Parity), "None"); // Nincs paritás
            _com.DataBits = 8; // 8 adatbit
            _com.StopBits = (StopBits)Enum.Parse(typeof(StopBits), "1"); // 1 stopbit
            _com.Handshake = (Handshake)Enum.Parse(typeof(Handshake), "None"); // Nincs kézfogás (flow control)
            _com.ReadTimeout = timeout; // Az időtúllépés olvasás esetén
            _com.WriteTimeout = timeout; // Az időtúllépés írás esetén

            _com.Open(); // A port megnyitása
        }

        /// <summary>Ez a függvény bezárja a soros portot</summary>
        /// <remarks>A függvény nem rendelkezik kivételkezeléssel, így a hívó függvénynek kell ezt megoldania.</remarks>
        public void CloseComPort()
        {
            _com.Close(); // A port bezárása
        }

        /// <summary>Ez a függvény elküld egy csomagot a porton keresztül, majd fogadja a választ</summary>
        /// <param name="packet">A küldendő csomag</param>
        /// <param name="receiveLength">A fogadanadó csomag hossza</param>
        /// <returns>A fogadott csomag</returns>
        /// <remarks>
        /// A függvény nem rendelkezik kivételkezeléssel, így a hívó függvénynek kell ezt megoldania.
        /// </remarks>
        public byte[] SendPacket(byte[] packet, ushort receiveLength)
        {
            if (receiveLength == 0) // Ha nem kell semmit sem fogadni,
                throw new ArgumentException(); // akkor kivételt dobunk

            byte[] receive = new byte[receiveLength]; // A fogadási buffer létrehozása, ide mentjük a kapott csomagot

            _com.Write(packet, 0, packet.GetLength(0)); // A csomag elküldése
            _com.DiscardInBuffer(); // Kitöröljük a hardveres fogadási buffert, hogy a fogadáskor a megfelelő bájtokat kapjuk

            for (int i = 0; i < receiveLength; i++) // Kiolvassuk az összes bájtot,
                receive[i] = (byte)_com.ReadByte();

            return receive; // és végül visszaadjuk a tömböt
        }

        /// <summary>Ez a függvény elküld egy csomagot a porton keresztül</summary>
        /// <param name="packet">A küldendő csomag</param>
        /// <remarks>
        /// A függvény nem rendelkezik kivételkezeléssel, így a hívó függvénynek kell ezt megoldania.
        /// </remarks>
        public void SendPacket(byte[] packet)
        {
            _com.Write(packet, 0, packet.GetLength(0)); // A csomag elküldése
            _com.DiscardInBuffer(); // Kitöröljük a hardveres fogadási buffert, hogy a fogadáskor a megfelelő bájtokat kapjuk
            return;
        }

        /// <summary>Ez a függvény fogad egy csomagot a porton keresztül</summary>
        /// <param name="receiveLength">A fogadanadó csomag hossza</param>
        /// <returns>A fogadott csomag</returns>
        /// <remarks>
        /// A függvény nem rendelkezik kivételkezeléssel, így a hívó függvénynek kell ezt megoldania.
        /// </remarks>
        public byte[] ReceivePacket(ushort receiveLength)
        {
            if (receiveLength == 0) // Ha nem kell semmit sem fogadni,
                throw new ArgumentException(); // akkor kivételt dobunk

            byte[] receive = new byte[receiveLength]; // A fogadási buffer létrehozása, ide mentjük a kapott csomagot

            for (int i = 0; i < receiveLength; i++) // Kiolvassuk az összes bájtot,
                receive[i] = (byte)_com.ReadByte();

            return receive; // és végül visszaadjuk a tömböt
        }

        /// <summary>Ez a függvény ellenőrzi, hogy a kapott csomag helyes-e</summary>
        /// <param name="incomingPacket">Az ellenőrizendő csomag</param>
        /// <param name="pidMustBe">A várt [PID] bájt</param>
        /// <param name="headerMustBe">A várt [HEADER] bájt</param>
        /// <returns>Igaz, ha a csomag helyes, hamis, ha érvénytelen</returns>
        /// <remarks>
        /// A csomag felépítése: [START][PID][ID][HEADER][DATA1]...[DATAn][CHK][STOP]
        /// Ez a függvény nem ellenőrzi az eszközazonosítót.
        /// </remarks>
        public bool IsPacketOk(byte[] incomingPacket, byte pidMustBe, byte headerMustBe)
        {
            byte chk = 0; // Az ellenőrző összeg kiszámolásához használt változó
            ushort packetLength = (ushort)incomingPacket.GetLength(0); // A csomag hossza

            if (incomingPacket[0] != Constants.slaveToMasterStart) return false; // A [START] bájt ellenőrzése
            if (incomingPacket[1] != pidMustBe) return false; // A [PID] bájt ellenőrzése
            if (incomingPacket[3] != headerMustBe) return false; // A [HEADER] bájt ellenőrzése

            for (ushort i = 2; i < (packetLength - 2); i++) // [CHK] = 0xFF - ([ID] + [HEADER] + [DATA1] + ... + [DATAn]) + 1
                chk += incomingPacket[i]; // Az ellenőrző összeg növelése az aktuális bájttal

            if (incomingPacket[packetLength - 2] != (byte)(0xFF - chk + 0x01)) return false; // A [CHK] bájt ellenőrzése
            if (incomingPacket[packetLength - 1] != Constants.slaveToMasterStop) return false; // A [STOP] bájt ellenőrzése

            return true; // A csomag rendben van
        }

        /// <summary>Ez a függvény ellenőrzi, hogy a kapott csomag helyes-e</summary>
        /// <param name="incomingPacket">Az ellenőrizendő csomag</param>
        /// <param name="pidMustBe">A várt [PID] bájt</param>
        /// <param name="idMustBe">A várt [ID] bájt</param>
        /// <param name="headerMustBe">A várt [HEADER] bájt</param>
        /// <returns>Igaz, ha a csomag helyes, hamis, ha érvénytelen</returns>
        /// <remarks>A csomag felépítése: [START][PID][ID][HEADER][DATA1]...[DATAn][CHK][STOP]</remarks>
        public bool IsPacketOk(byte[] incomingPacket, byte pidMustBe, byte idMustBe, byte headerMustBe)
        {
            byte chk = 0; // Az ellenőrző összeg kiszámolásához használt változó
            ushort packetLength = (ushort)incomingPacket.GetLength(0); // A csomag hossza

            if (incomingPacket[0] != Constants.slaveToMasterStart) return false; // A [START] bájt ellenőrzése
            if (incomingPacket[1] != pidMustBe) return false; // A [PID] bájt ellenőrzése
            if (incomingPacket[2] != idMustBe) return false; // A [ID] bájt ellenőrzése
            if (incomingPacket[3] != headerMustBe) return false; // A [HEADER] bájt ellenőrzése

            for (ushort i = 2; i < (packetLength - 2); i++) // [CHK] = 0xFF - ([ID] + [HEADER] + [DATA1] + ... + [DATAn]) + 1
                chk += incomingPacket[i]; // Az ellenőrző összeg növelése az aktuális bájttal

            if (incomingPacket[packetLength - 2] != (byte)(0xFF - chk + 0x01)) return false; // A [CHK] bájt ellenőrzése
            if (incomingPacket[packetLength - 1] != Constants.slaveToMasterStop) return false; // A [STOP] bájt ellenőrzése

            return true; // A csomag rendben van
        }

        /// <summary>Ez a függvény visszaadja egybájtos válaszcsomagok esetén a [DATA1] jelentését szövegként</summary>
        /// <param name="data1">A vizsgálandó [DATA1] bájt</param>
        /// <returns>A bájt jelentése szövegként</returns>
        public string DecodeAnswerData(byte data1)
        {
            switch (data1)
            {
                //
                // Az eszköz által küldött [DATA1] bájtjok jelentése
                //
                case Constants.pcAnswerOk: // A parancs sikeresen végre lett hajtva
                    return "The command has been successfully executed!";

                case Constants.pcAnswerError: // A parancs nem hajtható végre
                    return "The command can not be executed!";

                case Constants.pcAnswerErrorOutOfRange: // A parancs nem hajtható végre, mert az érték kívül esik a megengedett tartományon
                    return "The command can not be executed, because the value is out of the allowable range!";

                //
                // Nem létező [DATA1] bájtok
                //
                default:
                    return "Invalid [DATA1] byte!";
            }
        }

        /// <summary>
        /// Ez a függvény visszaadja egybájtos válaszcsomagok esetén a [DATA1] bájt jelentését arra vonatkozólag,
        /// hogy a kívánt beállítás sikeres volt-e
        /// </summary>
        /// <param name="data1">A vizsgálandó [DATA1] bájt</param>
        /// <returns>Igaz, ha a kért beállítás sikeresen alkalmazva lett</returns>
        public bool DecodeAnswerSuccess(byte data1)
        { // A parancs vagy most, vagy már korábban sikeresen végre lett hajtva
            if (data1 == Constants.pcAnswerOk) return true;
            else return false; // A parancs végrehajtása valamilyen okból kifolyólag sikertelen volt
        }
    }
}
