using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace universal_logger
{
    /// <summary>Ez az osztály tartalmazza az egységekkel való kommunikációhoz tartozó konstansokat</summary>
    class Constants
    {
        //
        // A mester és a szolga közti kommunikáció START, STOP bájtjai
        //
        /// <summary>A mester által a szolgának küldött csomagok START bájtja</summary>
        public const byte masterToSlaveStart = 0x3A;
        /// <summary>A mester által a szolgának küldött csomagok STOP bájtja</summary>
        public const byte masterToSlaveStop = 0x1B;
        /// <summary>A szolga által a mesternek küldött csomagok START bájtja</summary>
        public const byte slaveToMasterStart = 0xDE;
        /// <summary>A szolga által a mesternek küldött csomagok STOP bájtja</summary>
        public const byte slaveToMasterStop = 0xA3;

        //
        // A mester és a szolga közti kommunikáció fix HEADER bájtjai
        // [START][PID][ID][HEADER][DATA1]...[DATAn][CHK][STOP]
        //
        /// <summary>A szolga hívása</summary>
        public const byte headerCalling = 0x00;
        /// <summary>A szolga újraindítása</summary>
        public const byte headerReset = 0x01;
        /// <summary>A szolga típus- illetve működési információjának a lekérdezése</summary>
        public const byte headerAskType = 0x02;
        /// <summary>A szolga firmware verziószámának a lekérdezése</summary>
        public const byte headerAskFwVersion = 0x03;
        /// <summary>A szolga címének a beállítása</summary>
        public const byte headerSetAddress = 0x20;

        //
        // A szolga által a mesternek küldött egybájtos válaszcsomagok DATA1 bájtjai
        // [START][PID][ID][HEADER][DATA1][CHK][STOP]
        //
        /// <summary>[DATA1]: A parancs sikeresen végre lett hajtva</summary>
        public const byte pcAnswerOk = 0x00;
        /// <summary>[DATA1]: A parancs nem hajtható végre</summary>
        public const byte pcAnswerError = 0x10;
        /// <summary>[DATA1]: A parancs nem hajtható végre, mert az érték kívül esik a megengedett tartományon</summary>
        public const byte pcAnswerErrorOutOfRange = 0x11;
    }
}
