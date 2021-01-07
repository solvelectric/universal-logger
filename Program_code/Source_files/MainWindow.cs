using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace universal_logger
{
    /// <summary>Ez az osztály tartalmazza a fő programablakot, és ezzel együtt a főprogram elemeit is</summary>
    public partial class MainWindow : Form
    {
        /// <summary>A fájlműveletekért felelős objektum</summary>
        private TextFileWriter fileWriter;

        /// <summary>A paramétertáblázatot kezelő objektum</summary>
        private OperatingParameters opParam;

        /// <summary>A használt kommunikációs port</summary>
        private Communication comPort;

        /// <summary>Az eszköz adatait tartalmazó objektum</summary>
        private Device device;

        /// <summary>A lekérdezésekhez tartozó időzítő</summary>
        private Timer timerQuery;

        /// <summary>A napi e-mailek küldéséhez tartozó időzítő</summary>
        private Timer timerDailyEmail;

        /// <summary>Véletlenszám-generátor a csomagazonosítókhoz</summary>
        Random pidGenerator;

        /// <summary>Igaz, ha az időzítőknek már nem szabad lefutniuk</summary>
        private bool timersShouldStop;

        /// <summary>Ez a lista mondja meg, hogy a lekérdezhető adatok a táblázat mely sorában és oszlopában találhatóak</summary>
        /// <remarks>
        /// A legkülső lista a parancsok szerint van elosztva.
        /// A középső listában találhatóak az egyes adatok.
        /// A legbelső lista 0. eleme tartalmazza a sorindexet, az 1. eleme pedig az oszlopindexet.
        /// Például a 4. parancs 3. adatának a sorindexe queryDataLookupTable[3][2][0], oszlopindexe queryDataLookupTable[3][2][1].
        /// </remarks>
        private List<List<List<int>>> queryDataLookupTable;

        /// <summary>Ez a lista mondja meg, hogy a lekérdezhető bitek a táblázat mely sorában és oszlopában találhatóak</summary>
        /// <remarks>
        /// A legkülső lista a parancsok szerint van elosztva.
        /// Kívülről a második listában találhatóak az egyes adatok.
        /// Kívülről a harmadik listában találhatóak az egyes bitek.
        /// A legbelső lista 0. eleme tartalmazza a sorindexet, az 1. eleme pedig az oszlopindexet.
        /// Például a 1. parancs 4. adatában a 11. bit sorindexe queryBitsLookupTable[0][3][10][0],
        /// oszlopindexe: queryBitsLookupTable[0][3][10][1].
        /// </remarks>
        private List<List<List<List<int>>>> queryBitsLookupTable;

        /// <summary>Az egyes parancsokhoz tartozó kommunikációs hibák utolsó előfordulási idejét tartalmazó lista</summary>
        /// <remarks>
        /// Az index a parancs sorszáma.
        /// Például a 4. parancshoz tartozó adat a commFaultLastDate[3] helyen található.
        /// </remarks>
        private List<DateTime> commFaultLastDate;

        /// <summary>Igaz, ha az egyes parancsokhoz tartozó kommunikációs hibáknál már elküldtük az első e-mailt</summary>
        /// <remarks>
        /// Az index a parancs sorszáma.
        /// Például a 4. parancshoz tartozó adat a commFaultFirstEmailSent[3] helyen található.
        /// </remarks>
        private List<bool> commFaultFirstEmailSent;

        /// <summary>Igaz, ha az egyes parancsokhoz tartozó kommunikációs hibáknál már elküldtük a második e-mailt</summary>
        /// <remarks>
        /// Az index a parancs sorszáma.
        /// Például a 4. parancshoz tartozó adat a commFaultSecondEmailSent[3] helyen található.
        /// </remarks>
        private List<bool> commFaultSecondEmailSent;

        /// <summary>Az egyes bitekhez tartozó hibák utolsó előfordulásainak idejét tartalmazó lista</summary>
        /// <remarks>
        /// A legkülső lista a parancsok szerint van elosztva.
        /// Kívülről a második listában találhatóak az egyes adatok.
        /// Kívülről a harmadik listában találhatóak az egyes bitek.
        /// Például a 1. parancs 4. adatában a 11. bit sorindexe faultLastDate[0][3][10], oszlopindexe: faultLastDate[0][3][10].
        /// </remarks>
        private List<List<List<DateTime>>> faultLastDate;

        /// <summary>Igaz, ha az egyes bitekhez tartozó hibáknál már elküldtük az első e-mailt</summary>
        /// <remarks>
        /// A legkülső lista a parancsok szerint van elosztva.
        /// Kívülről a második listában találhatóak az egyes adatok.
        /// Kívülről a harmadik listában találhatóak az egyes bitek.
        /// Például a 1. parancs 4. adatában a 11. bit sorindexe faultFirstEmailSent[0][3][10],
        /// oszlopindexe: faultFirstEmailSent[0][3][10].
        /// </remarks>
        private List<List<List<bool>>> faultFirstEmailSent;

        /// <summary>Igaz, ha az egyes bitekhez tartozó hibáknál már elküldtük a második e-mailt</summary>
        /// <remarks>
        /// A legkülső lista a parancsok szerint van elosztva.
        /// Kívülről a második listában találhatóak az egyes adatok.
        /// Kívülről a harmadik listában találhatóak az egyes bitek.
        /// Például a 1. parancs 4. adatában a 11. bit sorindexe faultSecondEmailSent[0][3][10],
        /// oszlopindexe: faultSecondEmailSent[0][3][10].
        /// </remarks>
        private List<List<List<bool>>> faultSecondEmailSent;

        /// <summary>Ez a lista tartalmazza a beálltó parancsokhoz tartozó ablakokat</summary>
        private List<SetData> setDataWindows;

        /// <summary>A fő programablak inicializálása</summary>
        public MainWindow()
        {
            try
            {
                fileWriter = new TextFileWriter(Application.StartupPath); // A fájlkezelő objektum létrehozása
            }
            catch (Exception ex) // Ha kivétel történt a fájlok és a könyvtárak létrehozása során,
            {
                MessageBox.Show("An exception occured during the file/directory initialization:" + Environment.NewLine +
                                ex.Message + Environment.NewLine + // akkor megjelenítünk egy hibaüzenetet,
                                "The application will terminate.");

                this.Load += CloseOnStart; // és bezárjuk a programot
                return;
            }

            try
            {
                fileWriter.WriteExceptionFile(DateTime.Now, "Starting..."); // Loggoljuk a program indítását
            }
            catch (Exception ex) // ha nem sikerül az első fájlírás,
            {
                MessageBox.Show("An exception occured during the exception file initialization:" + Environment.NewLine +
                                ex.Message + Environment.NewLine + // akkor megjelenítünk egy hibaüzenetet,
                                "The application will terminate.");

                this.Load += CloseOnStart; // és bezárjuk a programot
                return;
            }

            SelectTable initDialog; // A paramétertáblázat-választó ablakhoz szükséges változók
            DialogResult result;

            try
            {
                initDialog = new SelectTable(Application.StartupPath); // A paramétertáblázat-választó ablak létrehozása,
                result = initDialog.ShowDialog(); // és megjelenítése

                if (initDialog.autoSelected) // A táblázat automatikusan lett kiválasztva
                    fileWriter.WriteExceptionFile(DateTime.Now, "Startup table automatically selected: " + initDialog.selectedTable); // Lementjük a választott táblát
                else // A táblázatot a felhasználó válaszotta ki
                    fileWriter.WriteExceptionFile(DateTime.Now, "Startup table: " + initDialog.selectedTable); // Lementjük a választott táblát
            }
            catch (Exception ex) // Ha kivétel történt,
            {
                DateTime date = DateTime.Now; // akkor lementjük az időt,
                MessageBox.Show("An exception occured during the initialization:" + Environment.NewLine +
                                ex.Message + Environment.NewLine + // majd megjelenítünk egy hibaüzenetet
                                "The application will terminate.");

                try
                {
                    fileWriter.WriteExceptionFile(date, "MainWindow initDialog: " + ex.Message); // A hiba kiírása a kivétel-fájlba
                }
                catch { } // Így is úgy is bezárul a program

                this.Load += CloseOnStart; // Bezárjuk a programot
                return;
            }

            InitializeComponent(); // Az ablak inicializálása

            if (result != DialogResult.OK) // Ha a felhasználó nem választott ki semmit,
            {
                this.Load += CloseOnStart; // akkor vége az alkalmazásnak
                return;
            }

            try // Ha a felhasználó kiválasztott egy paramétertáblázatot, akkor ellenőrizzük azt
            {
                opParam = new OperatingParameters(initDialog.selectedTable); // Létrehozzuk a paramétertáblazatot kezelő objektumot
                if (!opParam.Open()) // Ha a választott táblázat érvénytelen,
                {
                    MessageBox.Show("The choosen parameter table is incorrect!" + Environment.NewLine + // akkor megjelenítünk egy
                                    opParam.faultText + Environment.NewLine + // hibaüzenetet,
                                    "Please restart the application, and choose a valid table!" + Environment.NewLine +
                                    "If there is no valid parameter table, please contact with the system administrator!");

                    this.Load += CloseOnStart; // és bezárjuk a programot
                    return;
                }
            }
            catch (Exception ex) // Ha kivétel történt a paramétertáblázat megnyitása során,
            {
                DateTime date = DateTime.Now; // akkor lementjük az időt,
                MessageBox.Show("An exception occured during the opening of the parameter-table:" + Environment.NewLine +
                                ex.Message + Environment.NewLine + // majd megjelenítünk egy hibaüzenetet
                                "The application will terminate.");

                try
                {
                    fileWriter.WriteExceptionFile(date, "MainWindow opParam: " + ex.Message); // A hiba kiírása a kivétel-fájlba
                }
                catch { } // Így is úgy is bezárul a program

                this.Load += CloseOnStart; // Bezárjuk a programot
                return;
            }

            try // Megpróbáljuk létrehozni és megnyitni a COM portot
            { // A használt kommunikációs port
                comPort = new Communication(opParam.settingsData.port, opParam.settingsData.baudRate, opParam.settingsData.timeout);
            }
            catch (Exception ex) // Ha kivétel történt a port megnyitása során,
            {
                DateTime date = DateTime.Now; // akkor lementjük az időt,
                MessageBox.Show("An exception occured during the opening of COM port:" + Environment.NewLine +
                                ex.Message + Environment.NewLine + // majd megjelenítünk egy hibaüzenetet
                                "The application will terminate.");

                try
                {
                    fileWriter.WriteExceptionFile(date, "MainWindow comPort: " + ex.Message); // A hiba kiírása a kivétel-fájlba
                }
                catch { } // Így is úgy is bezárul a program

                this.Load += CloseOnStart; // Bezárjuk a programot
                return;
            }

            // A logfájlok könyvtárának a beállítása
            logDirectoryPathBox.Text = Application.StartupPath + System.IO.Path.DirectorySeparatorChar + opParam.settingsData.logDirectoryName;
            try
            { // A logfájl és a könyvtár létrehozása
                fileWriter.setLogFileParameters(logDirectoryPathBox.Text, opParam.settingsData.logFileHeader);
            }
            catch (Exception ex) // Ha kivétel történt a fájlok és a könyvtárak létrehozása során,
            {
                DateTime date = DateTime.Now; // akkor lementjük az időt,
                MessageBox.Show("An exception occured during the file/directory initialization:" + Environment.NewLine +
                                ex.Message + Environment.NewLine + // majd megjelenítünk egy hibaüzenetet
                                "The application will terminate.");

                try
                {
                    fileWriter.WriteExceptionFile(date, "MainWindow fileWriter: " + ex.Message); // A hiba kiírása a kivétel-fájlba
                }
                catch { } // Így is úgy is bezárul a program

                this.Load += CloseOnStart; // Bezárjuk a programot
                return;
            }

            device = new Device(0, opParam); // Létrehozzuk az eszközt leíró objektumot

            pidGenerator = new Random(); // A véletlenszám-generátor inicializálása

            // Inicializáljuk a lekérdezésekhez tartozó időzítést
            intervalBox.SelectedIndex = intervalBox.Items.IndexOf("5s"); // Az alapértelmezett időzítő intervallum 5s
            timerQuery = new Timer(); // Az időzítő inicializálása
            timerQuery.Interval = 5000; // Az időzítés beállítása
            timerQuery.Tick += new EventHandler(QueryTimerEvent); // Az időzítő jelzésekor lefutó függvény

            timerDailyEmail = new Timer(); // Inicializáljuk a napi e-mail-ek küldéséhez tartozó időzítést
            DateTime now = DateTime.Now;
            DateTime sendDate = now.Date.AddHours(opParam.settingsData.emailSendingHour); // A mai nap ekkor kellene küldeni az e-mailt
            if (sendDate < now) sendDate = sendDate.AddDays(1.0); // Ha ma már nem kell e-mail küldeni, akkor ugrunk a kövektező napra
            timerDailyEmail.Interval = (int)((sendDate - now).TotalMilliseconds); // Az időzítés beállítása
            timerDailyEmail.Tick += new EventHandler(DailyTimerEvent); // Az időzítő jelzésekor lefutó függvény
            timerDailyEmail.Start(); // Elindítjuk az időzítést

            timersShouldStop = false; // Az időzítők szabadon futhatnak

            int rowNumber; // Itt tároljuk le a táblázatok sorainak a számát

            rowNumber = (opParam.settingsData.numOfQueryCommands / 4); // A lekérdező parancsokat felsoroló táblázat sorainak a száma
            if ((opParam.settingsData.numOfQueryCommands % 4) != 0) rowNumber++; // Ha van "tört" sor, akkor 1-gyel növeljük a sorok számát
            if (rowNumber != 0) // Ha a lekérdező parancsokat tartalmazó táblázat tartalmaz sorokat,
                queryCommandsTable.Rows.Add(rowNumber); // akkor inicializáljuk őket
            for (byte i = 0; i < rowNumber; i++) // Végigmegyünk a sorokon
                queryCommandsTable.Rows[i].Height = 20; // és beállítjuk azok magasságát

            // A sorok számától függően növeljük a parancsokat tartalmazó rész méretét,
            queryCommandsGroup.Height += (rowNumber - 1) * 20; // eltoljuk az adatokat tartalmazó részt lefelé
            queryDataGroup.Location = new Point(queryDataGroup.Location.X, queryDataGroup.Location.Y + ((rowNumber - 1) * 20));
            queryDataGroup.Height -= (rowNumber - 1) * 20; // valamint ennek is csökkentjük a méretét

            for (byte i = 0; i < opParam.settingsData.numOfQueryCommands; i++) // Végigmegyünk a parancsokon, és feltöltjük a neveket
            { // a táblázatba, valamint a beállításoktól függően beállítjuk a jelölőnégyzeteket
                queryCommandsTable.Rows[i / 4].Cells[(2 * (i % 4)) + 1].Value = opParam.settingsData.queryCommands[i].name;
                if (opParam.settingsData.queryCommands[i].autoQuery) // Le kell kérdezni a parancsot
                    ((DataGridViewCheckBoxCell)(queryCommandsTable.Rows[i / 4].Cells[(2 * (i % 4))])).Value = true;
                else // Nem kell lekérdezni a parancsot
                    ((DataGridViewCheckBoxCell)(queryCommandsTable.Rows[i / 4].Cells[(2 * (i % 4))])).Value = false;
            }

            rowNumber = 0; // A lekérdezhető adatokat felsoroló táblázat sorainak a száma
            for (byte i = 0; i < opParam.settingsData.numOfQueryCommands; i++) // Végigmegyünk a parancsokon, és megszámoljuk
            { // az adatokhoz szükséges
                rowNumber += opParam.settingsData.queryCommands[i].numOfSections / 4; // sorok számát (parancsonként új sort kezdünk)
                if ((opParam.settingsData.queryCommands[i].numOfSections % 4) != 0) // Ha van "tört" sor, akkor
                    rowNumber++; // 1-gyel növeljük a sorok számát
            }
            if (rowNumber != 0) // Ha az adatokat megjelenítő táblázat tartalmaz sorokat,
                queryDataTable.Rows.Add(rowNumber); // akkor inicializáljuk őket
            for (ushort i = 0; i < rowNumber; i++) // Végigmegyünk a sorokon
                queryDataTable.Rows[i].Height = 20; // és beállítjuk azok magasságát

            // A lekérdezhető adatok elhelyezkedését tartalmazó lista létrehozása
            queryDataLookupTable = new List<List<List<int>>>(opParam.settingsData.numOfQueryCommands);

            int num = 0; // Itt tároljuk a következő adat sorszámát
            for (byte i = 0; i < opParam.settingsData.numOfQueryCommands; i++) // Végigmegyünk a lekérdezhető parancsokon
            { // Létrehozzuk a lekérdezhető adatok elhelyezkedését tartalmazó lista aktuális parancshoz tartozó elemét
                queryDataLookupTable.Add(new List<List<int>>(opParam.settingsData.queryCommands[i].numOfSections));
                if ((num % 4) != 0) num += (4 - (num % 4)); // Ha az előző sor tört sor volt, akkor ugrunk a következő sorra
                for (byte j = 0; j < opParam.settingsData.queryCommands[i].numOfSections; j++) // Végigmegyünk az adatokon,
                { // és feltöltjük a neveket a táblázatba
                    queryDataTable.Rows[num / 4].Cells[2 * (num % 4)].Value = opParam.settingsData.queryCommands[i].sections[j].name;
                    queryDataLookupTable[i].Add(new List<int>(2)); // Lementjük a lekérdezhető adatok elhelyezkedését tartalmazó lista
                    queryDataLookupTable[i][j].Add(num / 4); // aktuális adatához tartozó sor-
                    queryDataLookupTable[i][j].Add((2 * (num % 4)) + 1); // illetve oszlopindexet
                    num++; // A következő adat sorszáma
                }
            }

            rowNumber = 0; // A lekérdezhető biteket felsoroló táblázat sorainak a száma
            for (byte i = 0; i < opParam.settingsData.numOfQueryCommands; i++) // Végigmegyünk a parancsokon,
            { // és megszámoljuk a státusz bitekhez szükséges
                for (byte j = 0; j < opParam.settingsData.queryCommands[i].numOfSections; j++) // sorok számát (adatonként új sort kezdünk)
                {
                    byte numOfUsedBits = 0;
                    for (byte k = 0; k < opParam.settingsData.queryCommands[i].sections[j].numOfBits; k++) // Végigmegyünk a biteken,
                    {
                        if (opParam.settingsData.queryCommands[i].sections[j].bits[k].isUsed) // és megszámoljuk, hogy az adott bitek
                        numOfUsedBits++; // közül hány van használatban
                    }

                    rowNumber += numOfUsedBits / 4; // Növeljük a sorok számát
                    if ((numOfUsedBits % 4) != 0) // Ha van "tört" sor, akkor
                        rowNumber++; // 1-gyel növeljük a sorok számát
                }
            }
            if (rowNumber != 0) // Ha a biteket megjelenítő táblázat tartalmaz sorokat,
                statusBitsTable.Rows.Add(rowNumber); // akkor inicializáljuk őket
            for (int i = 0; i < rowNumber; i++) // Végigmegyünk a sorokon
                statusBitsTable.Rows[i].Height = 20; // és beállítjuk azok magasságát

            // A lekérdezhető bitek elhelyezkedését tartalmazó lista létrehozása
            queryBitsLookupTable = new List<List<List<List<int>>>>(opParam.settingsData.numOfQueryCommands);

            num = 0; // Itt tároljuk a következő bit sorszámát
            for (byte i = 0; i < opParam.settingsData.numOfQueryCommands; i++) // Végigmegyünk a lekérdezhető parancsokon
            { // Létrehozzuk a lekérdezhető bitek elhelyezkedését tartalmazó lista aktuális parancshoz tartozó elemét
                queryBitsLookupTable.Add(new List< List<List<int>>>(opParam.settingsData.queryCommands[i].numOfSections));
                for (byte j = 0; j < opParam.settingsData.queryCommands[i].numOfSections; j++) // Végigmegyünk a lekérdezhető adatokon
                { // Létrehozzuk a lekérdezhető bitek elhelyezkedését tartalmazó lista aktuális adatához tartozó elemét
                    queryBitsLookupTable[i].Add(new List<List<int>>(opParam.settingsData.queryCommands[i].sections[j].numOfBits));
                    if ((num % 4) != 0) num += (4 - (num % 4)); // Ha az előző sor tört sor volt, akkor ugrunk a következő sorra
                    for (byte k = 0; k < opParam.settingsData.queryCommands[i].sections[j].numOfBits; k++) // Végigmegyünk a biteken
                    { // Létrehozzuk a lekérdezhető bitek elhelyezkedését tartalmazó lista aktuális bitjéhez tartozó elemét
                        queryBitsLookupTable[i][j].Add(new List<int>(2));
                        if (opParam.settingsData.queryCommands[i].sections[j].bits[k].isUsed) // Ha az adott bit használatban van,
                        {
                            statusBitsTable.Rows[num / 4].Cells[(2 * (num % 4)) + 1].Value = // akkor feltöltjük a neveket a táblázatba,
                                opParam.settingsData.queryCommands[i].sections[j].bits[k].name;
                            queryBitsLookupTable[i][j][k].Add(num / 4); // valamint lementjük az adott bithez tartozó sor-
                            queryBitsLookupTable[i][j][k].Add(2 * (num % 4)); // és oszlopindexet
                            num++; // A következő adat sorszáma
                        }
                        else // Ha az adott bit nincs használatban,
                        {
                            queryBitsLookupTable[i][j][k].Add(-1); // akkor ezt -1-es sor- illetve oszlopindexxel jelezzük
                            queryBitsLookupTable[i][j][k].Add(-1); // az adott bithez tartozó listaelemnél
                        }
                    }
                }
            }

            rowNumber = opParam.settingsData.numOfSetCommands / 4; // A beállító parancsokat felsoroló táblázat sorainak a száma
            if ((opParam.settingsData.numOfSetCommands % 4) != 0) rowNumber++; // Ha van "tört" sor, akkor 1-gyel növeljük a sorok számát
            if (rowNumber != 0) // Ha a beállító parancsokat tartalmazó táblázat tartalmaz sorokat,
                setCommandsTable.Rows.Add(rowNumber); // akkor inicializáljuk őket
            for (byte i = 0; i < rowNumber; i++) // Végigmegyünk a sorokon
                setCommandsTable.Rows[i].Height = 20; // és beállítjuk azok magasságát

            for (byte i = 0; i < opParam.settingsData.numOfSetCommands; i++) // Végigmegyünk a parancsokon, és feltöltjük a neveket
                setCommandsTable.Rows[i / 4].Cells[i % 4].Value = opParam.settingsData.setCommands[i].name; // a táblázatba

            // A hibákat, és azok előfordulásának adatait tartalmazó listák létrehozása, valamint a listák elemeinek feltöltése
            commFaultLastDate = new List<DateTime>(opParam.settingsData.numOfQueryCommands);
            commFaultFirstEmailSent = new List<bool>(opParam.settingsData.numOfQueryCommands);
            commFaultSecondEmailSent = new List<bool>(opParam.settingsData.numOfQueryCommands);
            faultLastDate = new List<List<List<DateTime>>>(opParam.settingsData.numOfQueryCommands);
            faultFirstEmailSent = new List<List<List<bool>>>(opParam.settingsData.numOfQueryCommands);
            faultSecondEmailSent = new List<List<List<bool>>>(opParam.settingsData.numOfQueryCommands);
            for (byte i = 0; i < opParam.settingsData.numOfQueryCommands; i++) // Végigmegyünk a parancsokon
            {
                commFaultLastDate.Add(new DateTime(0)); // Kezdetben még nem küldtünk egy e-mailt sem
                commFaultFirstEmailSent.Add(false);
                commFaultSecondEmailSent.Add(false);
                faultLastDate.Add(new List<List<DateTime>>(opParam.settingsData.queryCommands[i].numOfSections));
                faultFirstEmailSent.Add(new List<List<bool>>(opParam.settingsData.queryCommands[i].numOfSections));
                faultSecondEmailSent.Add(new List<List<bool>>(opParam.settingsData.queryCommands[i].numOfSections));
                for (byte j = 0; j < opParam.settingsData.queryCommands[i].numOfSections; j++) // Végigmegyünk az adatokon
                {
                    faultLastDate[i].Add(new List<DateTime>(opParam.settingsData.queryCommands[i].sections[j].numOfBits));
                    faultFirstEmailSent[i].Add(new List<bool>(opParam.settingsData.queryCommands[i].sections[j].numOfBits));
                    faultSecondEmailSent[i].Add(new List<bool>(opParam.settingsData.queryCommands[i].sections[j].numOfBits));
                    for (byte k = 0; k < opParam.settingsData.queryCommands[i].sections[j].numOfBits; k++) // Végigmegyünk a biteken
                    {
                        faultLastDate[i][j].Add(new DateTime(0)); // Kezdetben még nem küldtünk egy e-mailt sem
                        faultFirstEmailSent[i][j].Add(false);
                        faultSecondEmailSent[i][j].Add(false);
                    }
                }
            }

            // Létrehozzuk a beállító parancsokhoz tartozó ablakok listáját,
            setDataWindows = new List<SetData>(opParam.settingsData.numOfSetCommands);
            for (byte i = 0; i < opParam.settingsData.numOfSetCommands; i++) // majd feltöltjük elemekkel
                setDataWindows.Add(new SetData(i, opParam));

            idBox.Text = "1"; // Beállítjuk az 1-es című eszközt
            idApplyButton_Click(this, EventArgs.Empty);

            // Végezetül az automatikus lekérdezést is elindítjuk, ha szükséges
            if (opParam.settingsData.timerAutoStart != 0) // Ha a paramétertáblában megadtak egy intervallumot, akkor kiválasztjuk a
            {
                intervalBox.SelectedIndex = opParam.settingsData.timerAutoStart - 1; // paramétertáblázatban megjelölt időintervallumot,
                timerButton_Click(this, EventArgs.Empty); // majd elindítjuk az időzítőt
            }

            return;
        }

        /// <summary>Ez a függvény zárja be az inicializálás során a programot</summary>
        private void CloseOnStart(object sender, EventArgs e)
        {
            try
            {
                fileWriter.WriteExceptionFile(DateTime.Now, "Closing..."); // Loggoljuk a program bezárását
            }
            catch { } // Így is úgy is bezárul a program

            this.Dispose(); // A program bezárása
            return;
        }

        /// <summary>Az ablak bezárásakor létrejövő esemény</summary>
        /// <param name="e">Az esemény paraméterei</param>
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            try
            {
                if (e.CloseReason == CloseReason.UserClosing) // Ha a felhasználó akarja bezárni a programot,
                { // akkor megkérdezzük, hogy biztos így akar-e tenni
                    DialogResult result = MessageBox.Show("Are you sure you want to quit?", "Confirmation", MessageBoxButtons.YesNo);
                    if (result != DialogResult.Yes) // Ha a felhasználó nem akarja bezárni az ablakot,
                    {
                        e.Cancel = true; // akkor nem megyünk tovább
                        return;
                    }
                }
            }
            catch { } // Így is úgy is bezárul a program

            try
            {
                lock (timerQuery)
                {
                    timersShouldStop = true; // Jelezzük, hogy az időzítő már nem futhat le
                }
            }
            catch { } // Így is úgy is bezárul a program

            try
            {
                fileWriter.WriteExceptionFile(DateTime.Now, "Closing..."); // Loggoljuk a program bezárását
            }
            catch { } // Így is úgy is bezárul a program

            base.OnFormClosing(e); // majd folytatjuk a program bezárását
        }
        
        /// <summary>A napi e-mail küldését figyelő időzítő jelzésekor lefutó függvény</summary>
        /// <param name="sender">Az eseményt küldő objektum</param>
        /// <param name="e">Az esemény paraméterei</param>
        private void DailyTimerEvent(object sender, EventArgs e)
        {
            try
            {
                lock (timerQuery)
                {
                    if (timersShouldStop) return; // Ha már nem futhatnak az időzítők, akkor nem csinálunk semmit sem
                    // Holnap ekkor kellene küldeni az e-mailt
                    DateTime sendDate = DateTime.Now.Date.AddDays(1).AddHours(opParam.settingsData.emailSendingHour); // Az időzítés
                    timerDailyEmail.Interval = (int)((sendDate - DateTime.Now).TotalMilliseconds); // beállítása, majd az e-mail elküldése
                    fileWriter.SendDailyMail(DateTime.Now, this.Text, opParam.settingsData.systemName, opParam.settingsData.emailList);
                }
                return;
            }
            catch (Exception ex) // Ha kivétel történt,
            {
                if (timersShouldStop) return; // Ha már nem futhatnak az időzítők, akkor nem csinálunk semmit sem

                DateTime date = DateTime.Now; // akkor lementjük az időt,

                try
                {
                    fileWriter.SendExceptionMail(date, this.Text, opParam.settingsData.systemName,
                        opParam.settingsData.emailList, ex); // elküldünk egy e-mailt a felhasználónak,
                    MessageBox.Show("An exception occured:" + Environment.NewLine + ex.Message); // majd megjelenítünk egy hibaüzenetet
                }
                catch // Ha nem tudtuk elküldeni az e-mailt, akkor kiírjuk a hibát a kivétel-fájlba
                {
                    try
                    {
                        fileWriter.WriteExceptionFile(date, "DailyTimerEvent_2: " + ex.Message); // A hiba kiírása a kivétel-fájlba
                    }
                    catch (Exception exc) // Ha valamiért nem tudunk írni a kivétel-fájlba, akkor bezárjuk a programot
                    {
                        MessageBox.Show("An exception occured during file writing:" + Environment.NewLine +
                                        exc.Message + Environment.NewLine + // akkor megjelenítünk egy hibaüzenetet
                                        "The application will terminate.");

                        this.Close(); // A program bezárása
                    }
                }

                try
                {
                    fileWriter.WriteExceptionFile(date, "DailyTimerEvent_1: " + ex.Message); // A hiba kiírása a kivétel-fájlba
                }
                catch (Exception exc) // Ha valamiért nem tudunk írni a kivétel-fájlba, akkor bezárjuk a programot
                {
                    MessageBox.Show("An exception occured during file writing:" + Environment.NewLine +
                                    exc.Message + Environment.NewLine + // akkor megjelenítünk egy hibaüzenetet
                                    "The application will terminate.");

                    this.Close(); // A program bezárása
                }

                timerQuery.Start(); // Újraindítjuk az időzítőt
                return;
            }
        }

        /// <summary>Az lekérdezést figyelő időzítő jelzésekor lefutó függvény</summary>
        /// <param name="sender">Az eseményt küldő objektum</param>
        /// <param name="e">Az esemény paraméterei</param>
        private void QueryTimerEvent(object sender, EventArgs e)
        {
            try
            {
                timerQuery.Stop(); // Megállítjuk az időzítőt,
                lock (timerQuery) // majd letiltjuk a többi műveletet amíg nem végzünk itt
                {
                    if (timersShouldStop) return; // Ha már nem futhatnak az időzítők, akkor nem csinálunk semmit sem
                    string description = ""; // Amennyiben történt valamilyen hiba, akkor azt itt tároljuk le
                    DateTime now = DateTime.Now; // Lementjük az időt, hogy minden hiba esetén ugyanaz az érték szerepeljen

                    for (byte index = 0; index < opParam.settingsData.numOfQueryCommands; index++) // Végigmegyünk a lekérdező parancsokon
                    {
                        DataGridViewCheckBoxCell cell = // Megvizsgáljuk az adott lekérdező parancshoz tartozó jelölőnégyzetet
                            (DataGridViewCheckBoxCell)queryCommandsTable.Rows[index / 4].Cells[(2 * (index % 4))];

                        if (!((bool)(cell.Value))) continue; // Ha a jelölőnégyzet nincs bepipálva, akkor ugrunk a következő parancsra

                        // A jelölőnégyzet be van pipálva, tehát a parancsot végre kell hajtani

                        if (device.AskGeneral(comPort, (byte)pidGenerator.Next(1, 256), (byte)index, opParam))
                        {
                            // Ha a lekérdezés sikeres volt, akkor végigmegyünk a parancshoz tartozó adatokon, és frissítjük őket a
                            for (byte i = 0; i < opParam.settingsData.queryCommands[index].numOfSections; i++) // felhasználói felületen
                            {
                                if (opParam.settingsData.queryCommands[index].sections[i].decimals == -2) // Az adat hexadecimális
                                {
                                    // Lementjük az adat típusát, amely csak előjel nélküli egész lehet
                                    Type typeTmp = opParam.settingsData.queryCommands[index].sections[i].type;
                                    if (typeTmp == typeof(byte)) // Az adat 8 bites előjel nélküli egész
                                    {
                                        queryDataTable.Rows[queryDataLookupTable[index][i][0]]. // Lementjük az adatot
                                            Cells[queryDataLookupTable[index][i][1]].Value = // hexadecimális formátumban
                                            ((byte)(device.queryData[index][i])).ToString("X2");
                                    }
                                    else if (typeTmp == typeof(ushort)) // Az adat 16 bites előjel nélküli egész
                                    {
                                        queryDataTable.Rows[queryDataLookupTable[index][i][0]]. // Lementjük az adatot
                                            Cells[queryDataLookupTable[index][i][1]].Value = // hexadecimális formátumban
                                            ((ushort)(device.queryData[index][i])).ToString("X4");
                                    }
                                    else // Az adat 32 bites előjel nélküli egész
                                    {
                                        queryDataTable.Rows[queryDataLookupTable[index][i][0]]. // Lementjük az adatot
                                            Cells[queryDataLookupTable[index][i][1]].Value = // hexadecimális formátumban
                                            ((uint)(device.queryData[index][i])).ToString("X8");
                                    }

                                    // Az adatokat tartalmazó táblázatban már frissítettük az értéket, így a bitek következnek
                                    // Végigmegyünk a biteken
                                    for (byte j = 0; j < opParam.settingsData.queryCommands[index].sections[i].numOfBits; j++)
                                    {
                                        if (queryBitsLookupTable[index][i][j][0] != -1) // Ha az adott bit használatban van,
                                        {
                                            // akkor megvizsgáljuk a típusát
                                            if (typeTmp == typeof(byte)) // Az adat 8 bites előjel nélküli egész
                                            {
                                                if ((((byte)(device.queryData[index][i])) & ((byte)(1 << (7 - j)))) != 0) // Ha a bit
                                                { // logikai 1, akkor bepipáljuk a jelölőnégyzetet,
                                                    ((DataGridViewCheckBoxCell)(statusBitsTable.Rows[queryBitsLookupTable[index][i][j][0]].
                                                        Cells[queryBitsLookupTable[index][i][j][1]])).Value = true;

                                                    if (opParam.settingsData.queryCommands[index].sections[i].bits[j].sendEmail)
                                                    { // és ha szükséges, akkor pirosra állítjuk a cella hátterét
                                                        statusBitsTable.Rows[queryBitsLookupTable[index][i][j][0]].
                                                            Cells[queryBitsLookupTable[index][i][j][1] + 1].Style.BackColor =
                                                            Color.FromArgb(255, 75, 75);

                                                        if (faultLastDate[index][i][j].AddHours(1.0) <= now) // Ha már legalább
                                                        { // 1 óra eltelt az utolsó e-mail óta, valamint még nem küldtük el
                                                            if (!faultFirstEmailSent[index][i][j]) // az első e-mailt erről a hibáról,
                                                            {
                                                                description += "Fault: " + // akkor inicializáljuk a hiba szövegét,
                                                                    opParam.settingsData.queryCommands[index].sections[i].bits[j].name +
                                                                    Environment.NewLine; // frissítjük a küldés idejét,
                                                                faultLastDate[index][i][j] = now; // majd jelezzük,
                                                                faultFirstEmailSent[index][i][j] = true; // hogy mostmár elküldtük az e-mailt
                                                            }
                                                            else if (!faultSecondEmailSent[index][i][j]) // Ha a hiba továbbra is fennáll,
                                                            { // de még csak egy e-mailt küldtünk erről a hibáról,
                                                                description += "Fault: " + // akkor inicializáljuk a hiba szövegét,
                                                                    opParam.settingsData.queryCommands[index].sections[i].bits[j].name +
                                                                    Environment.NewLine + "You won't receive another e-mail about this " +
                                                                    "fault, until it won't disappear." + Environment.NewLine; // frissítjük
                                                                faultLastDate[index][i][j] = now; // a küldés idejét, majd jelezzük,
                                                                faultSecondEmailSent[index][i][j] = true; // hogy elküldtük mindkét e-mailt
                                                            }
                                                            // Ha már két e-mailt is küldtünk a hibáról úgy, hogy az közben nem szűnt meg,
                                                            // akkor nem küldünk több e-mailt
                                                        }
                                                        // Egy órán belül nem küldünk e-mailt ugyanazon hibáról akkor sem,
                                                        // ha oszcillál a hiba
                                                    }
                                                }
                                                else // Ha a bit logikai 0,
                                                { // akkor töröljük a pipát a jelölőnégyzetből,
                                                    ((DataGridViewCheckBoxCell)(statusBitsTable.Rows[queryBitsLookupTable[index][i][j][0]].
                                                        Cells[queryBitsLookupTable[index][i][j][1]])).Value = false;

                                                    if (opParam.settingsData.queryCommands[index].sections[i].bits[j].sendEmail)
                                                    { // és ha szükséges, akkor töröljük a cella hátterét
                                                        statusBitsTable.Rows[queryBitsLookupTable[index][i][j][0]].
                                                            Cells[queryBitsLookupTable[index][i][j][1] + 1].Style.BackColor = Color.Empty;

                                                        faultFirstEmailSent[index][i][j] = false; // Jelezzük, hogy nem küldtünk
                                                        faultSecondEmailSent[index][i][j] = false; // még e-mailt a hibáról
                                                    }
                                                }
                                            }
                                            else if (typeTmp == typeof(ushort)) // Az adat 16 bites előjel nélküli egész
                                            {
                                                if ((((ushort)(device.queryData[index][i])) & ((ushort)(1 << (15 - j)))) != 0) // Ha a bit
                                                { // logikai 1, akkor bepipáljuk a jelölőnégyzetet,
                                                    ((DataGridViewCheckBoxCell)(statusBitsTable.Rows[queryBitsLookupTable[index][i][j][0]].
                                                        Cells[queryBitsLookupTable[index][i][j][1]])).Value = true;

                                                    if (opParam.settingsData.queryCommands[index].sections[i].bits[j].sendEmail)
                                                    { // és ha szükséges, akkor pirosra állítjuk a cella hátterét
                                                        statusBitsTable.Rows[queryBitsLookupTable[index][i][j][0]].
                                                            Cells[queryBitsLookupTable[index][i][j][1] + 1].Style.BackColor =
                                                            Color.FromArgb(255, 75, 75);

                                                        if (faultLastDate[index][i][j].AddHours(1.0) <= now) // Ha már legalább
                                                        { // 1 óra eltelt az utolsó e-mail óta, valamint még nem küldtük el
                                                            if (!faultFirstEmailSent[index][i][j]) // az első e-mailt erről a hibáról,
                                                            {
                                                                description += "Fault: " + // akkor inicializáljuk a hiba szövegét,
                                                                    opParam.settingsData.queryCommands[index].sections[i].bits[j].name +
                                                                    Environment.NewLine; // frissítjük a küldés idejét,
                                                                faultLastDate[index][i][j] = now; // majd jelezzük,
                                                                faultFirstEmailSent[index][i][j] = true; // hogy mostmár elküldtük az e-mailt
                                                            }
                                                            else if (!faultSecondEmailSent[index][i][j]) // Ha a hiba továbbra is fennáll,
                                                            { // de még csak egy e-mailt küldtünk erről a hibáról,
                                                                description += "Fault: " + // akkor inicializáljuk a hiba szövegét,
                                                                    opParam.settingsData.queryCommands[index].sections[i].bits[j].name +
                                                                    Environment.NewLine + "You won't receive another e-mail about this " +
                                                                    "fault, until it won't disappear." + Environment.NewLine; // frissítjük
                                                                faultLastDate[index][i][j] = now; // a küldés idejét, majd jelezzük,
                                                                faultSecondEmailSent[index][i][j] = true; // hogy elküldtük mindkét e-mailt
                                                            }
                                                            // Ha már két e-mailt is küldtünk a hibáról úgy, hogy az közben nem szűnt meg,
                                                            // akkor nem küldünk több e-mailt
                                                        }
                                                        // Egy órán belül nem küldünk e-mailt ugyanazon hibáról akkor sem,
                                                        // ha oszcillál a hiba
                                                    }
                                                }
                                                else // Ha a bit logikai 0,
                                                { // akkor töröljük a pipát a jelölőnégyzetből,
                                                    ((DataGridViewCheckBoxCell)(statusBitsTable.Rows[queryBitsLookupTable[index][i][j][0]].
                                                        Cells[queryBitsLookupTable[index][i][j][1]])).Value = false;

                                                    if (opParam.settingsData.queryCommands[index].sections[i].bits[j].sendEmail)
                                                    { // és ha szükséges, akkor töröljük a cella hátterét
                                                        statusBitsTable.Rows[queryBitsLookupTable[index][i][j][0]].
                                                            Cells[queryBitsLookupTable[index][i][j][1] + 1].Style.BackColor = Color.Empty;

                                                        faultFirstEmailSent[index][i][j] = false; // Jelezzük, hogy nem küldtünk
                                                        faultSecondEmailSent[index][i][j] = false; // még e-mailt a hibáról
                                                    }
                                                }
                                            }
                                            else // Az adat 32 bites előjel nélküli egész
                                            {
                                                if ((((uint)(device.queryData[index][i])) & ((uint)(1 << (31 - j)))) != 0) // Ha a bit
                                                { // logikai 1, akkor bepipáljuk a jelölőnégyzetet,
                                                    ((DataGridViewCheckBoxCell)(statusBitsTable.Rows[queryBitsLookupTable[index][i][j][0]].
                                                        Cells[queryBitsLookupTable[index][i][j][1]])).Value = true;

                                                    if (opParam.settingsData.queryCommands[index].sections[i].bits[j].sendEmail)
                                                    { // és ha szükséges, akkor pirosra állítjuk a cella hátterét
                                                        statusBitsTable.Rows[queryBitsLookupTable[index][i][j][0]].
                                                            Cells[queryBitsLookupTable[index][i][j][1] + 1].Style.BackColor =
                                                            Color.FromArgb(255, 75, 75);

                                                        if (faultLastDate[index][i][j].AddHours(1.0) <= now) // Ha már legalább
                                                        { // 1 óra eltelt az utolsó e-mail óta, valamint még nem küldtük el
                                                            if (!faultFirstEmailSent[index][i][j]) // az első e-mailt erről a hibáról,
                                                            {
                                                                description += "Fault: " + // akkor inicializáljuk a hiba szövegét,
                                                                    opParam.settingsData.queryCommands[index].sections[i].bits[j].name +
                                                                    Environment.NewLine; // frissítjük a küldés idejét,
                                                                faultLastDate[index][i][j] = now; // majd jelezzük,
                                                                faultFirstEmailSent[index][i][j] = true; // hogy mostmár elküldtük az e-mailt
                                                            }
                                                            else if (!faultSecondEmailSent[index][i][j]) // Ha a hiba továbbra is fennáll,
                                                            { // de még csak egy e-mailt küldtünk erről a hibáról,
                                                                description += "Fault: " + // akkor inicializáljuk a hiba szövegét,
                                                                    opParam.settingsData.queryCommands[index].sections[i].bits[j].name +
                                                                    Environment.NewLine + "You won't receive another e-mail about this " +
                                                                    "fault, until it won't disappear." + Environment.NewLine; // frissítjük
                                                                faultLastDate[index][i][j] = now; // a küldés idejét, majd jelezzük,
                                                                faultSecondEmailSent[index][i][j] = true; // hogy elküldtük mindkét e-mailt
                                                            }
                                                            // Ha már két e-mailt is küldtünk a hibáról úgy, hogy az közben nem szűnt meg,
                                                            // akkor nem küldünk több e-mailt
                                                        }
                                                        // Egy órán belül nem küldünk e-mailt ugyanazon hibáról akkor sem,
                                                        // ha oszcillál a hiba
                                                    }
                                                }
                                                else // Ha a bit logikai 0,
                                                { // akkor töröljük a pipát a jelölőnégyzetből,
                                                    ((DataGridViewCheckBoxCell)(statusBitsTable.Rows[queryBitsLookupTable[index][i][j][0]].
                                                        Cells[queryBitsLookupTable[index][i][j][1]])).Value = false;

                                                    if (opParam.settingsData.queryCommands[index].sections[i].bits[j].sendEmail)
                                                    { // és ha szükséges, akkor töröljük a cella hátterét
                                                        statusBitsTable.Rows[queryBitsLookupTable[index][i][j][0]].
                                                            Cells[queryBitsLookupTable[index][i][j][1] + 1].Style.BackColor = Color.Empty;

                                                        faultFirstEmailSent[index][i][j] = false; // Jelezzük, hogy nem küldtünk
                                                        faultSecondEmailSent[index][i][j] = false; // még e-mailt a hibáról
                                                    }
                                                }
                                            }
                                        }
                                        // Ha az adott bit nincs használatban, akkor nincs semmilyen teendőnk sem
                                    }
                                    // Miután az összes bitet megvizsgáltuk, nincs több teendőnk
                                }
                                else if (opParam.settingsData.queryCommands[index].sections[i].decimals == -1) // Az adat dátum/idő
                                {
                                    queryDataTable.Rows[queryDataLookupTable[index][i][0]]. // Lementjük az adatot
                                        Cells[queryDataLookupTable[index][i][1]].Value =
                                        ((DateTime)(device.queryData[index][i])).ToString("yyyy.MM.dd HH\\:mm\\:ss");
                                }
                                else // Az adat egy "normál" (nem hexadecimális) szám (egész vagy tört)
                                {
                                    float value; // A szám értéke
                                    string format = "F"; // valamint formátuma

                                    value = Convert.ToSingle(device.queryData[index][i]); // A szám értéke egészként

                                    for (byte count = 0; count < opParam.settingsData.queryCommands[index].sections[i].decimals; count++)
                                    { // Ahány tizedesjegyet tartalmaz az érték,
                                        value /= 10.0F; // annyiszor osztjuk le tízzel
                                    }

                                    // Kiegészítjük a formátum sztringet a tizedesjegyek számával
                                    format += opParam.settingsData.queryCommands[index].sections[i].decimals.ToString();

                                    queryDataTable.Rows[queryDataLookupTable[index][i][0]]. // Lementjük az adatot a megfelelő formátumban
                                        Cells[queryDataLookupTable[index][i][1]].Value = value.ToString(format);
                                }
                            }
                            // Végeztünk az összes, ehhez a parancshoz tartozó adattal

                            commFaultFirstEmailSent[index] = false; // Jelezzük, hogy még nem küldtünk e-mailt a kommunikációs hibáról
                            commFaultSecondEmailSent[index] = false;
                        }
                        else // A kommunikáció nem sikerült
                        {
                            if (commFaultLastDate[index].AddHours(1.0) <= now) // Ha már legalább 1 óra eltelt az utolsó e-mail óta,
                            {
                                if (!commFaultFirstEmailSent[index]) // valamint még nem küldtük el az első e-mailt erről a hibáról,
                                {
                                    description += "Communication fault at the following command: " + // akkor inicializáljuk a
                                        opParam.settingsData.queryCommands[index].name + Environment.NewLine; // hiba szövegét,
                                    commFaultLastDate[index] = now; // frissítjük a küldés idejét,
                                    commFaultFirstEmailSent[index] = true; // majd jelezzük, hogy mostmár elküldtük az e-mailt
                                }
                                else if (!commFaultSecondEmailSent[index]) // Ha a hiba továbbra is fennáll,
                                { // de még csak egy e-mailt küldtünk erről a hibáról,
                                    description += "Communication fault at the following command: " + // akkor inicializáljuk a
                                        opParam.settingsData.queryCommands[index].name + Environment.NewLine + // hiba szövegét,
                                        "You won't receive another e-mail about this fault, until it won't disappear." + Environment.NewLine;
                                    commFaultLastDate[index] = now; // frissítjük a küldés idejét,
                                    commFaultSecondEmailSent[index] = true; // majd jelezzük, elküldtük mindkét e-mailt
                                }
                                // Ha már két e-mailt is küldtünk a hibáról úgy, hogy az közben nem szűnt meg, akkor nem küldünk több e-mailt
                            }
                            // Egy órán belül nem küldünk e-mailt ugyanazon hibáról akkor sem, ha oszcillál a hiba
                        }
                    }

                    // Minden lekérdező paranccsal végeztünk

                    if (description != "") // Ha történt valamilyen hiba, amiről értesíteni kell a felhasználót,
                        fileWriter.SendFaultMail(now, this.Text, opParam.settingsData.systemName, opParam.settingsData.emailList,
                            description); // akkor elküldjük az e-mailt

                    if (!opParam.settingsData.useDeviceDateTime) // Ha nem a műszer szolgáltatja a lekérdezés idejét,
                        device.currentDateTime = now; // akkor lementjük a mostani időt

                    fileWriter.WriteLogFile(device.currentDateTime, device.CreateLogFileLine(opParam)); // Kiírjuk az adatokat a log fájlba

                    dataReadBox.Text = device.lastReceivedData; // Az állapotsor frissítése (csak az utolsó értékeket írjuk ki)
                    messageReadBox.Text = device.lastReceivedDataDecoded;
                }

                timerQuery.Start(); // Újraindítjuk az időzítőt
                return;
            }
            catch (Exception ex) // Ha kivétel történt,
            {
                if (timersShouldStop) return; // Ha már nem futhatnak az időzítők, akkor nem csinálunk semmit sem

                DateTime date = DateTime.Now; // akkor lementjük az időt,

                try
                {
                    fileWriter.SendExceptionMail(date, this.Text, opParam.settingsData.systemName,
                        opParam.settingsData.emailList, ex); // elküldünk egy e-mailt a felhasználónak,
                    MessageBox.Show("An exception occured:" + Environment.NewLine + ex.Message); // majd megjelenítünk egy hibaüzenetet
                }
                catch // Ha nem tudtuk elküldeni az e-mailt, akkor kiírjuk a hibát a kivétel-fájlba
                {
                    try
                    {
                        fileWriter.WriteExceptionFile(date, "QueryTimerEvent_2: " + ex.Message); // A hiba kiírása a kivétel-fájlba
                    }
                    catch (Exception exc) // Ha valamiért nem tudunk írni a kivétel-fájlba, akkor bezárjuk a programot
                    {
                        MessageBox.Show("An exception occured during file writing:" + Environment.NewLine +
                                        exc.Message + Environment.NewLine + // akkor megjelenítünk egy hibaüzenetet
                                        "The application will terminate.");

                        this.Close(); // A program bezárása
                    }
                }

                try
                {
                    fileWriter.WriteExceptionFile(date, "QueryTimerEvent_1: " + ex.Message); // A hiba kiírása a kivétel-fájlba
                }
                catch (Exception exc) // Ha valamiért nem tudunk írni a kivétel-fájlba, akkor bezárjuk a programot
                {
                    MessageBox.Show("An exception occured during file writing:" + Environment.NewLine +
                                    exc.Message + Environment.NewLine + // akkor megjelenítünk egy hibaüzenetet
                                    "The application will terminate.");

                    this.Close(); // A program bezárása
                }

                timerQuery.Start(); // Újraindítjuk az időzítőt
                return;
            }
        }

        /// <summary>A használt eszközazonosító a kiválasztása</summary>
        /// <param name="sender">Az eseményt küldő objektum</param>
        /// <param name="e">Az esemény paraméterei</param>
        private void idApplyButton_Click(object sender, EventArgs e)
        {
            try
            {
                lock (timerQuery) // Letiltjuk a többi műveletet amíg nem végzünk itt
                {
                    try
                    {
                        byte tmpUi8 = byte.Parse(idBox.Text); // Beolvassuk a számot,
                        device.id = tmpUi8; // majd beállítjuk,
                        idReadBox.Text = device.id.ToString(); // és átmásoljuk a másik szövegdobozba
                    }
                    catch (OverflowException) // Ha a szám túl kicsi, vagy túl nagy,
                    {
                        MessageBox.Show("0 <= ID <= 255"); // akkor megjelenítjük a hibaüzenetet
                        return;
                    }
                }

                return;
            }
            catch (Exception ex) // Ha kivétel történt,
            {
                DateTime date = DateTime.Now; // akkor lementjük az időt,
                MessageBox.Show("An exception occured:" + Environment.NewLine + ex.Message); // majd megjelenítünk egy hibaüzenetet

                try
                {
                    fileWriter.WriteExceptionFile(date, "idApplyButton_Click: " + ex.Message); // A hiba kiírása a kivétel-fájlba
                }
                catch (Exception exc) // Ha valamiért nem tudunk írni a kivétel-fájlba, akkor bezárjuk a programot
                {
                    MessageBox.Show("An exception occured during file writing:" + Environment.NewLine +
                                    exc.Message + Environment.NewLine + // akkor megjelenítünk egy hibaüzenetet
                                    "The application will terminate.");

                    this.Close(); // A program bezárása
                }

                return;
            }
        }

        /// <summary>Az eszköz eszközazonosítójának a megváltoztatása</summary>
        /// <param name="sender">Az eseményt küldő objektum</param>
        /// <param name="e">Az esemény paraméterei</param>
        private void newIdSetButton_Click(object sender, EventArgs e)
        {
            try
            {
                lock (timerQuery) // Letiltjuk a többi műveletet amíg nem végzünk itt
                {
                    byte tmpUi8 = 0;

                    try
                    {
                        tmpUi8 = byte.Parse(newIdBox.Text); // Beolvassuk a számot
                    }
                    catch (OverflowException) // Ha a szám túl kicsi, vagy túl nagy,
                    {
                        MessageBox.Show("0 <= ID <= 255"); // akkor megjelenítjük a hibaüzenetet
                        return;
                    }

                    if (device.SendNewId(comPort, (byte)pidGenerator.Next(1, 256), tmpUi8)) // Elküldjük az új eszközazonosítót,
                    { // és ha a parancs végrehajtása sikeres volt,
                        idReadBox.Text = device.id.ToString(); // akkor beírjuk az új értéket szövegdobozba
                    }

                    dataReadBox.Text = device.lastReceivedData; // Az állapotsor frissítése
                    messageReadBox.Text = device.lastReceivedDataDecoded;
                }

                return;
            }
            catch (Exception ex) // Ha kivétel történt,
            {
                DateTime date = DateTime.Now; // akkor lementjük az időt,
                MessageBox.Show("An exception occured:" + Environment.NewLine + ex.Message); // majd megjelenítünk egy hibaüzenetet

                try
                {
                    fileWriter.WriteExceptionFile(date, "newIdSetButton_Click: " + ex.Message); // A hiba kiírása a kivétel-fájlba
                }
                catch (Exception exc) // Ha valamiért nem tudunk írni a kivétel-fájlba, akkor bezárjuk a programot
                {
                    MessageBox.Show("An exception occured during file writing:" + Environment.NewLine +
                                    exc.Message + Environment.NewLine + // akkor megjelenítünk egy hibaüzenetet
                                    "The application will terminate.");

                    this.Close(); // A program bezárása
                }

                return;
            }
        }

        /// <summary>Az eszköz hívása</summary>
        /// <param name="sender">Az eseményt küldő objektum</param>
        /// <param name="e">Az esemény paraméterei</param>
        private void callingButton_Click(object sender, EventArgs e)
        {
            try
            {
                lock (timerQuery) // Letiltjuk a többi műveletet amíg nem végzünk itt
                {
                    device.Calling(comPort, (byte)pidGenerator.Next(1, 256)); // Az eszköz hívása

                    dataReadBox.Text = device.lastReceivedData; // Az állapotsor frissítése
                    messageReadBox.Text = device.lastReceivedDataDecoded;
                }

                return;
            }
            catch (Exception ex) // Ha kivétel történt,
            {
                DateTime date = DateTime.Now; // akkor lementjük az időt,
                MessageBox.Show("An exception occured:" + Environment.NewLine + ex.Message); // majd megjelenítünk egy hibaüzenetet

                try
                {
                    fileWriter.WriteExceptionFile(date, "callingButton_Click: " + ex.Message); // A hiba kiírása a kivétel-fájlba
                }
                catch (Exception exc) // Ha valamiért nem tudunk írni a kivétel-fájlba, akkor bezárjuk a programot
                {
                    MessageBox.Show("An exception occured during file writing:" + Environment.NewLine +
                                    exc.Message + Environment.NewLine + // akkor megjelenítünk egy hibaüzenetet
                                    "The application will terminate.");

                    this.Close(); // A program bezárása
                }

                return;
            }
        }

        /// <summary>Az eszköz újraindítása</summary>
        /// <param name="sender">Az eseményt küldő objektum</param>
        /// <param name="e">Az esemény paraméterei</param>
        private void resetButton_Click(object sender, EventArgs e)
        {
            try
            {
                lock (timerQuery) // Letiltjuk a többi műveletet amíg nem végzünk itt
                {
                    device.Reset(comPort, (byte)pidGenerator.Next(1, 256)); // Az eszköz újraindítása

                    dataReadBox.Text = device.lastReceivedData; // Az állapotsor frissítése
                    messageReadBox.Text = device.lastReceivedDataDecoded;
                }

                return;
            }
            catch (Exception ex) // Ha kivétel történt,
            {
                DateTime date = DateTime.Now; // akkor lementjük az időt,
                MessageBox.Show("An exception occured:" + Environment.NewLine + ex.Message); // majd megjelenítünk egy hibaüzenetet

                try
                {
                    fileWriter.WriteExceptionFile(date, "resetButton_Click: " + ex.Message); // A hiba kiírása a kivétel-fájlba
                }
                catch (Exception exc) // Ha valamiért nem tudunk írni a kivétel-fájlba, akkor bezárjuk a programot
                {
                    MessageBox.Show("An exception occured during file writing:" + Environment.NewLine +
                                    exc.Message + Environment.NewLine + // akkor megjelenítünk egy hibaüzenetet
                                    "The application will terminate.");

                    this.Close(); // A program bezárása
                }

                return;
            }
        }

        /// <summary>Az eszköz firmware verziószámának a lekérdezése</summary>
        /// <param name="sender">Az eseményt küldő objektum</param>
        /// <param name="e">Az esemény paraméterei</param>
        private void askVersionButton_Click(object sender, EventArgs e)
        {
            try
            {
                lock (timerQuery) // Letiltjuk a többi műveletet amíg nem végzünk itt
                {
                    if (device.AskFirmwareVersion(comPort, (byte)pidGenerator.Next(1, 256)))
                    { // Ha sikeresen le tudtuk kérdezni az értéket,
                        versionReadBox.Text = device.firmwareVersion.ToString(); // akkor megjelenítjük azt a kezelőfelületn
                    }

                    dataReadBox.Text = device.lastReceivedData; // Az állapotsor frissítése
                    messageReadBox.Text = device.lastReceivedDataDecoded;
                }

                return;
            }
            catch (Exception ex) // Ha kivétel történt,
            {
                DateTime date = DateTime.Now; // akkor lementjük az időt,
                MessageBox.Show("An exception occured:" + Environment.NewLine + ex.Message); // majd megjelenítünk egy hibaüzenetet

                try
                {
                    fileWriter.WriteExceptionFile(date, "askVersionButton_Click: " + ex.Message); // A hiba kiírása a kivétel-fájlba
                }
                catch (Exception exc) // Ha valamiért nem tudunk írni a kivétel-fájlba, akkor bezárjuk a programot
                {
                    MessageBox.Show("An exception occured during file writing:" + Environment.NewLine +
                                    exc.Message + Environment.NewLine + // akkor megjelenítünk egy hibaüzenetet
                                    "The application will terminate.");

                    this.Close(); // A program bezárása
                }

                return;
            }
        }

        /// <summary>Az eszköz típus- illetve működési információjának a lekérdezése</summary>
        /// <param name="sender">Az eseményt küldő objektum</param>
        /// <param name="e">Az esemény paraméterei</param>
        private void askTypeButton_Click(object sender, EventArgs e)
        {
            try
            {
                lock (timerQuery) // Letiltjuk a többi műveletet amíg nem végzünk itt
                {
                    if (device.AskType(comPort, (byte)pidGenerator.Next(1, 256)))
                    { // Ha sikeresen le tudtuk kérdezni az értéket,
                        typeReadBox.Text = device.type.ToString(); // akkor megjelenítjük azt a kezelőfelületn
                    }

                    dataReadBox.Text = device.lastReceivedData; // Az állapotsor frissítése
                    messageReadBox.Text = device.lastReceivedDataDecoded;
                }

                return;
            }
            catch (Exception ex) // Ha kivétel történt,
            {
                DateTime date = DateTime.Now; // akkor lementjük az időt,
                MessageBox.Show("An exception occured:" + Environment.NewLine + ex.Message); // majd megjelenítünk egy hibaüzenetet

                try
                {
                    fileWriter.WriteExceptionFile(date, "askTypeButton_Click: " + ex.Message); // A hiba kiírása a kivétel-fájlba
                }
                catch (Exception exc) // Ha valamiért nem tudunk írni a kivétel-fájlba, akkor bezárjuk a programot
                {
                    MessageBox.Show("An exception occured during file writing:" + Environment.NewLine +
                                    exc.Message + Environment.NewLine + // akkor megjelenítünk egy hibaüzenetet
                                    "The application will terminate.");

                    this.Close(); // A program bezárása
                }

                return;
            }
        }

        /// <summary>A felhasználó elindította vagy leállította az adatgyűjtést</summary>
        /// <param name="sender">Az eseményt küldő objektum</param>
        /// <param name="e">Az esemény paraméterei</param>
        private void timerButton_Click(object sender, EventArgs e)
        {
            try
            {
                if (timerButton.Text == "Start timer") // Most akarják elindítani az időzítést
                {
                    string intervalText = intervalBox.SelectedItem.ToString(); // Lekérjük a kiválasztott intervallumot
                    int intervalInt;

                    switch (intervalText) // "Átalakítjuk" a szöveget számmá
                    {
                        case "1s": intervalInt = 1000; break;
                        case "2s": intervalInt = 2000; break;
                        case "5s": intervalInt = 5000; break;
                        case "10s": intervalInt = 10000; break;
                        case "20s": intervalInt = 20000; break;
                        case "30s": intervalInt = 30000; break;
                        case "1min": intervalInt = 60000; break;
                        case "2min": intervalInt = 120000; break;
                        case "5min": intervalInt = 300000; break;
                        case "10min": intervalInt = 600000; break;
                        case "30min": intervalInt = 1800000; break;
                        case "60min": intervalInt = 3600000; break;
                        default: intervalInt = 5000; break; // Hiba esetén az alapértelmezett érték 5s
                    }

                    timerQuery.Interval = intervalInt; // Beállítjuk a felhasználó által kért intervallumot
                    timerQuery.Start(); // Elindítjuk az időzítést
                    timerButton.Text = "Stop timer";

                    intervalBox.Enabled = false; // Letiltjuk az intervallumválasztó lenyíló menüt
                }
                else // Most akarják megállítani az időzítést
                {
                    timerQuery.Stop(); // Megállítjuk az időzítést
                    timerButton.Text = "Start timer";

                    intervalBox.Enabled = true; // Engedélyezzük az intervallumválasztó lenyíló menüt
                }

                return;
            }
            catch (Exception ex) // Ha kivétel történt,
            {
                DateTime date = DateTime.Now; // akkor lementjük az időt,
                MessageBox.Show("An exception occured:" + Environment.NewLine + ex.Message); // majd megjelenítünk egy hibaüzenetet

                try
                {
                    fileWriter.WriteExceptionFile(date, "timerButton_Click: " + ex.Message); // A hiba kiírása a kivétel-fájlba
                }
                catch (Exception exc) // Ha valamiért nem tudunk írni a kivétel-fájlba, akkor bezárjuk a programot
                {
                    MessageBox.Show("An exception occured during file writing:" + Environment.NewLine +
                                    exc.Message + Environment.NewLine + // akkor megjelenítünk egy hibaüzenetet
                                    "The application will terminate.");

                    this.Close(); // A program bezárása
                }

                return;
            }
        }

        /// <summary>A logfájlok elérési útvonalának megváltoztatása</summary>
        /// <param name="sender">Az eseményt küldő objektum</param>
        /// <param name="e">Az esemény paraméterei</param>
        private void changeLogFilePath_Click(object sender, EventArgs e)
        {
            try
            {
                lock (timerQuery) // Letiltjuk a többi műveletet amíg nem végzünk itt
                {
                    FolderBrowserDialog dlg = new FolderBrowserDialog(); // A könyvtár-választó ablak létrehozása
                    dlg.Description = "Select the directory of the log files!";
                    dlg.SelectedPath = logDirectoryPathBox.Text; // A kezdő könyvtár a jelenleg kiválasztott könyvtár

                    if (dlg.ShowDialog() == DialogResult.OK) // Ha a felhasználó kiválasztott egy könyvtárat,
                    {
                        logDirectoryPathBox.Text = dlg.SelectedPath; // akkor beírjuk ezt a szövegdobozba, és létrehozzuk az új logfájlt
                        fileWriter.setLogFileParameters(logDirectoryPathBox.Text, opParam.settingsData.logFileHeader);
                    }
                }

                return;
            }
            catch (Exception ex) // Ha kivétel történt,
            {
                DateTime date = DateTime.Now; // akkor lementjük az időt,
                MessageBox.Show("An exception occured:" + Environment.NewLine + ex.Message); // majd megjelenítünk egy hibaüzenetet

                try
                {
                    fileWriter.WriteExceptionFile(date, "changeLogFilePath_Click: " + ex.Message); // A hiba kiírása a kivétel-fájlba
                }
                catch (Exception exc) // Ha valamiért nem tudunk írni a kivétel-fájlba, akkor bezárjuk a programot
                {
                    MessageBox.Show("An exception occured during file writing:" + Environment.NewLine +
                                    exc.Message + Environment.NewLine + // akkor megjelenítünk egy hibaüzenetet
                                    "The application will terminate.");

                    this.Close(); // A program bezárása
                }

                return;
            }
        }

        /// <summary>Akkor fut le, ha kijelöltek egy cellát abban a táblázatban, ahol a lekérdező parancsokat kiválasztani</summary>
        /// <param name="sender">Az eseményt küldő objektum</param>
        /// <param name="e">Az esemény paraméterei</param>
        private void queryCommandsTable_SelectionChanged(object sender, EventArgs e)
        {
            try
            {
                if (queryCommandsTable.SelectedCells.Count > 0) // Ha kijelöltek legalább egy cellát,
                    queryCommandsTable.SelectedCells[0].Selected = false; // akkor megszüntetjük a kijelölést
                return;
            }
            catch (Exception ex) // Ha kivétel történt,
            {
                DateTime date = DateTime.Now; // akkor lementjük az időt,
                MessageBox.Show("An exception occured:" + Environment.NewLine + ex.Message); // majd megjelenítünk egy hibaüzenetet

                try
                { // A hiba kiírása a kivétel-fájlba
                    fileWriter.WriteExceptionFile(date, "queryCommandsTable_SelectionChanged: " + ex.Message);
                }
                catch (Exception exc) // Ha valamiért nem tudunk írni a kivétel-fájlba, akkor bezárjuk a programot
                {
                    MessageBox.Show("An exception occured during file writing:" + Environment.NewLine +
                                    exc.Message + Environment.NewLine + // akkor megjelenítünk egy hibaüzenetet
                                    "The application will terminate.");

                    this.Close(); // A program bezárása
                }

                return;
            }
        }

        /// <summary>A felhasználó rákattinott egy cellára abban a táblázatban, ahol a lekérdező parancsokat kiválasztani</summary>
        /// <param name="sender">Az eseményt küldő objektum</param>
        /// <param name="e">Az esemény paraméterei</param>
        private void queryCommandsTable_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                int index = (e.RowIndex * 8 + e.ColumnIndex) / 2; // Az oszlop- és sorindexből kiszámoljuk a parancs indexét

                if (opParam.settingsData.numOfQueryCommands <= index) // Ha nem használt cellára kattintottak,
                    return; // akkor nem csinálunk semmit sem

                lock (timerQuery) // Letiltjuk a többi műveletet amíg nem végzünk itt
                {
                    if ((e.ColumnIndex % 2) == 1) // A gombra kattintottak, azaz végre kell hajtani a lekérdezést
                    {
                        if (device.AskGeneral(comPort, (byte)pidGenerator.Next(1, 256), (byte)index, opParam))
                        {
                            // Ha a lekérdezés sikeres volt, akkor végigmegyünk a parancshoz tartozó adatokon, és frissítjük őket a
                            for (byte i = 0; i < opParam.settingsData.queryCommands[index].numOfSections; i++) // felhasználói felületen
                            {
                                if (opParam.settingsData.queryCommands[index].sections[i].decimals == -2) // Az adat hexadecimális
                                {
                                    // Lementjük az adat típusát, amely csak előjel nélküli egész lehet
                                    Type typeTmp = opParam.settingsData.queryCommands[index].sections[i].type;
                                    if (typeTmp == typeof(byte)) // Az adat 8 bites előjel nélküli egész
                                    {
                                        queryDataTable.Rows[queryDataLookupTable[index][i][0]]. // Lementjük az adatot
                                            Cells[queryDataLookupTable[index][i][1]].Value = // hexadecimális formátumban
                                            ((byte)(device.queryData[index][i])).ToString("X2");
                                    }
                                    else if (typeTmp == typeof(ushort)) // Az adat 16 bites előjel nélküli egész
                                    {
                                        queryDataTable.Rows[queryDataLookupTable[index][i][0]]. // Lementjük az adatot
                                            Cells[queryDataLookupTable[index][i][1]].Value = // hexadecimális formátumban
                                            ((ushort)(device.queryData[index][i])).ToString("X4");
                                    }
                                    else // Az adat 32 bites előjel nélküli egész
                                    {
                                        queryDataTable.Rows[queryDataLookupTable[index][i][0]]. // Lementjük az adatot
                                            Cells[queryDataLookupTable[index][i][1]].Value = // hexadecimális formátumban
                                            ((uint)(device.queryData[index][i])).ToString("X8");
                                    }

                                    // Az adatokat tartalmazó táblázatban már frissítettük az értéket, így a bitek következnek
                                    // Végigmegyünk a biteken
                                    for (byte j = 0; j < opParam.settingsData.queryCommands[index].sections[i].numOfBits; j++)
                                    {
                                        if (queryBitsLookupTable[index][i][j][0] != -1) // Ha az adott bit használatban van,
                                        {
                                            // akkor megvizsgáljuk a típusát
                                            if (typeTmp == typeof(byte)) // Az adat 8 bites előjel nélküli egész
                                            {
                                                if ((((byte)(device.queryData[index][i])) & ((byte)(1 << (7 - j)))) != 0) // Ha a bit
                                                { // logikai 1, akkor bepipáljuk a jelölőnégyzetet,
                                                    ((DataGridViewCheckBoxCell)(statusBitsTable.Rows[queryBitsLookupTable[index][i][j][0]].
                                                        Cells[queryBitsLookupTable[index][i][j][1]])).Value = true;

                                                    if (opParam.settingsData.queryCommands[index].sections[i].bits[j].sendEmail)
                                                    { // és ha szükséges, akkor pirosra állítjuk a cella hátterét
                                                        statusBitsTable.Rows[queryBitsLookupTable[index][i][j][0]].
                                                            Cells[queryBitsLookupTable[index][i][j][1] + 1].Style.BackColor =
                                                            Color.FromArgb(255, 75, 75);
                                                    }
                                                }
                                                else // Ha a bit logikai 0,
                                                { // akkor töröljük a pipát a jelölőnégyzetből,
                                                    ((DataGridViewCheckBoxCell)(statusBitsTable.Rows[queryBitsLookupTable[index][i][j][0]].
                                                        Cells[queryBitsLookupTable[index][i][j][1]])).Value = false;

                                                    if (opParam.settingsData.queryCommands[index].sections[i].bits[j].sendEmail)
                                                    { // és ha szükséges, akkor töröljük a cella hátterét
                                                        statusBitsTable.Rows[queryBitsLookupTable[index][i][j][0]].
                                                            Cells[queryBitsLookupTable[index][i][j][1] + 1].Style.BackColor = Color.Empty;
                                                    }
                                                }
                                            }
                                            else if (typeTmp == typeof(ushort)) // Az adat 16 bites előjel nélküli egész
                                            {
                                                if ((((ushort)(device.queryData[index][i])) & ((ushort)(1 << (15 - j)))) != 0) // Ha a bit
                                                { // logikai 1, akkor bepipáljuk a jelölőnégyzetet,
                                                    ((DataGridViewCheckBoxCell)(statusBitsTable.Rows[queryBitsLookupTable[index][i][j][0]].
                                                        Cells[queryBitsLookupTable[index][i][j][1]])).Value = true;

                                                    if (opParam.settingsData.queryCommands[index].sections[i].bits[j].sendEmail)
                                                    { // és ha szükséges, akkor pirosra állítjuk a cella hátterét
                                                        statusBitsTable.Rows[queryBitsLookupTable[index][i][j][0]].
                                                            Cells[queryBitsLookupTable[index][i][j][1] + 1].Style.BackColor =
                                                            Color.FromArgb(255, 75, 75);
                                                    }
                                                }
                                                else // Ha a bit logikai 0,
                                                { // akkor töröljük a pipát a jelölőnégyzetből,
                                                    ((DataGridViewCheckBoxCell)(statusBitsTable.Rows[queryBitsLookupTable[index][i][j][0]].
                                                        Cells[queryBitsLookupTable[index][i][j][1]])).Value = false;

                                                    if (opParam.settingsData.queryCommands[index].sections[i].bits[j].sendEmail)
                                                    { // és ha szükséges, akkor töröljük a cella hátterét
                                                        statusBitsTable.Rows[queryBitsLookupTable[index][i][j][0]].
                                                            Cells[queryBitsLookupTable[index][i][j][1] + 1].Style.BackColor = Color.Empty;
                                                    }
                                                }
                                            }
                                            else // Az adat 32 bites előjel nélküli egész
                                            {
                                                if ((((uint)(device.queryData[index][i])) & ((uint)(1 << (31 - j)))) != 0) // Ha a bit
                                                { // logikai 1, akkor bepipáljuk a jelölőnégyzetet,
                                                    ((DataGridViewCheckBoxCell)(statusBitsTable.Rows[queryBitsLookupTable[index][i][j][0]].
                                                        Cells[queryBitsLookupTable[index][i][j][1]])).Value = true;

                                                    if (opParam.settingsData.queryCommands[index].sections[i].bits[j].sendEmail)
                                                    { // és ha szükséges, akkor pirosra állítjuk a cella hátterét
                                                        statusBitsTable.Rows[queryBitsLookupTable[index][i][j][0]].
                                                            Cells[queryBitsLookupTable[index][i][j][1] + 1].Style.BackColor =
                                                            Color.FromArgb(255, 75, 75);
                                                    }
                                                }
                                                else // Ha a bit logikai 0,
                                                { // akkor töröljük a pipát a jelölőnégyzetből,
                                                    ((DataGridViewCheckBoxCell)(statusBitsTable.Rows[queryBitsLookupTable[index][i][j][0]].
                                                        Cells[queryBitsLookupTable[index][i][j][1]])).Value = false;

                                                    if (opParam.settingsData.queryCommands[index].sections[i].bits[j].sendEmail)
                                                    { // és ha szükséges, akkor töröljük a cella hátterét
                                                        statusBitsTable.Rows[queryBitsLookupTable[index][i][j][0]].
                                                            Cells[queryBitsLookupTable[index][i][j][1] + 1].Style.BackColor = Color.Empty;
                                                    }
                                                }
                                            }
                                        }
                                        // Ha az adott bit nincs használatban, akkor nincs semmilyen teendőnk sem
                                    }
                                    // Miután az összes bitet megvizsgáltuk, nincs több teendőnk
                                }
                                else if (opParam.settingsData.queryCommands[index].sections[i].decimals == -1) // Az adat dátum/idő
                                {
                                    queryDataTable.Rows[queryDataLookupTable[index][i][0]]. // Lementjük az adatot
                                        Cells[queryDataLookupTable[index][i][1]].Value = // hexadecimális formátumban
                                        ((DateTime)(device.queryData[index][i])).ToString("yyyy.MM.dd HH\\:mm\\:ss");
                                }
                                else // Az adat egy "normál" (nem hexadecimális) szám (egész vagy tört)
                                {
                                    float value; // A szám értéke
                                    string format = "F"; // valamint formátuma
                                    
                                    value = Convert.ToSingle(device.queryData[index][i]); // A szám értéke egészként
                                    
                                    for (byte count = 0; count < opParam.settingsData.queryCommands[index].sections[i].decimals; count++)
                                    { // Ahány tizedesjegyet tartalmaz az érték,
                                        value /= 10.0F; // annyiszor osztjuk le tízzel
                                    }

                                    // Kiegészítjük a formátum sztringet a tizedesjegyek számával
                                    format += opParam.settingsData.queryCommands[index].sections[i].decimals.ToString();

                                    queryDataTable.Rows[queryDataLookupTable[index][i][0]]. // Lementjük az adatot a megfelelő formátumban
                                        Cells[queryDataLookupTable[index][i][1]].Value = value.ToString(format);
                                }
                            }
                            // Végeztünk az összes, ehhez a parancshoz tartozó adattal
                        }
                        // Ha a lekérdezés nem járt sikerrel, akkor nincs más teendőnk

                        dataReadBox.Text = device.lastReceivedData; // Az állapotsor frissítése
                        messageReadBox.Text = device.lastReceivedDataDecoded;
                    }
                    else // A jelölőnégyzetre kattintottak, azaz változtatni kell annak állapotát
                    {
                        DataGridViewCheckBoxCell cell = (DataGridViewCheckBoxCell)queryCommandsTable.Rows[e.RowIndex].Cells[e.ColumnIndex];

                        if ((bool)(cell.Value)) cell.Value = false; // Ha a jelölőnégyzet be van pipálva, akkor töröljük a pipát
                        else cell.Value = true; // ellenkező esetben pedig töröljük
                    }
                }

                return;
            }
            catch (Exception ex) // Ha kivétel történt,
            {
                DateTime date = DateTime.Now; // akkor lementjük az időt,
                MessageBox.Show("An exception occured:" + Environment.NewLine + ex.Message); // majd megjelenítünk egy hibaüzenetet

                try
                {
                    fileWriter.WriteExceptionFile(date, "queryCommandsTable_CellClick: " + ex.Message); // A hiba kiírása a kivétel-fájlba
                }
                catch (Exception exc) // Ha valamiért nem tudunk írni a kivétel-fájlba, akkor bezárjuk a programot
                {
                    MessageBox.Show("An exception occured during file writing:" + Environment.NewLine +
                                    exc.Message + Environment.NewLine + // akkor megjelenítünk egy hibaüzenetet
                                    "The application will terminate.");

                    this.Close(); // A program bezárása
                }

                return;
            }
        }

        /// <summary>Akkor fut le, ha kijelöltek egy cellát abban a táblázatban, ahol a lekérdezhető adatok vannak</summary>
        /// <param name="sender">Az eseményt küldő objektum</param>
        /// <param name="e">Az esemény paraméterei</param>
        private void queryDataTable_SelectionChanged(object sender, EventArgs e)
        {
            try
            {
                if (queryDataTable.SelectedCells.Count > 0) // Ha kijelöltek legalább egy cellát,
                    queryDataTable.SelectedCells[0].Selected = false; // akkor megszüntetjük a kijelölést
                return;
            }
            catch (Exception ex) // Ha kivétel történt,
            {
                DateTime date = DateTime.Now; // akkor lementjük az időt,
                MessageBox.Show("An exception occured:" + Environment.NewLine + ex.Message); // majd megjelenítünk egy hibaüzenetet

                try
                { // A hiba kiírása a kivétel-fájlba
                    fileWriter.WriteExceptionFile(date, "queryDataTable_SelectionChanged: " + ex.Message);
                }
                catch (Exception exc) // Ha valamiért nem tudunk írni a kivétel-fájlba, akkor bezárjuk a programot
                {
                    MessageBox.Show("An exception occured during file writing:" + Environment.NewLine +
                                    exc.Message + Environment.NewLine + // akkor megjelenítünk egy hibaüzenetet
                                    "The application will terminate.");

                    this.Close(); // A program bezárása
                }

                return;
            }
        }

        /// <summary>Akkor fut le, ha kijelöltek egy cellát abban a táblázatban, ahol a lekérdezhető bitek vannak</summary>
        /// <param name="sender">Az eseményt küldő objektum</param>
        /// <param name="e">Az esemény paraméterei</param>
        private void statusBitsTable_SelectionChanged(object sender, EventArgs e)
        {
            try
            {
                if (statusBitsTable.SelectedCells.Count > 0) // Ha kijelöltek legalább egy cellát,
                    statusBitsTable.SelectedCells[0].Selected = false; // akkor megszüntetjük a kijelölést
                return;
            }
            catch (Exception ex) // Ha kivétel történt,
            {
                DateTime date = DateTime.Now; // akkor lementjük az időt,
                MessageBox.Show("An exception occured:" + Environment.NewLine + ex.Message); // majd megjelenítünk egy hibaüzenetet

                try
                { // A hiba kiírása a kivétel-fájlba
                    fileWriter.WriteExceptionFile(date, "statusBitsTable_SelectionChanged: " + ex.Message);
                }
                catch (Exception exc) // Ha valamiért nem tudunk írni a kivétel-fájlba, akkor bezárjuk a programot
                {
                    MessageBox.Show("An exception occured during file writing:" + Environment.NewLine +
                                    exc.Message + Environment.NewLine + // akkor megjelenítünk egy hibaüzenetet
                                    "The application will terminate.");

                    this.Close(); // A program bezárása
                }

                return;
            }
        }

        /// <summary>Akkor fut le, ha kijelöltek egy cellát abban a táblázatban, ahol a beállító parancsokat kiválasztani</summary>
        /// <param name="sender">Az eseményt küldő objektum</param>
        /// <param name="e">Az esemény paraméterei</param>
        private void setCommandsTable_SelectionChanged(object sender, EventArgs e)
        {
            try
            {
                if (setCommandsTable.SelectedCells.Count > 0) // Ha kijelöltek legalább egy cellát,
                    setCommandsTable.SelectedCells[0].Selected = false; // akkor megszüntetjük a kijelölést
                return;
            }
            catch (Exception ex) // Ha kivétel történt,
            {
                DateTime date = DateTime.Now; // akkor lementjük az időt,
                MessageBox.Show("An exception occured:" + Environment.NewLine + ex.Message); // majd megjelenítünk egy hibaüzenetet

                try
                { // A hiba kiírása a kivétel-fájlba
                    fileWriter.WriteExceptionFile(date, "setCommandsTable_SelectionChanged: " + ex.Message);
                }
                catch (Exception exc) // Ha valamiért nem tudunk írni a kivétel-fájlba, akkor bezárjuk a programot
                {
                    MessageBox.Show("An exception occured during file writing:" + Environment.NewLine +
                                    exc.Message + Environment.NewLine + // akkor megjelenítünk egy hibaüzenetet
                                    "The application will terminate.");

                    this.Close(); // A program bezárása
                }

                return;
            }
        }

        /// <summary>A felhasználó rákattinott egy cellára abban a táblázatban, ahol a beállító parancsokat kiválasztani</summary>
        /// <param name="sender">Az eseményt küldő objektum</param>
        /// <param name="e">Az esemény paraméterei</param>
        private void setCommandsTable_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                int index = e.RowIndex * 4 + e.ColumnIndex; // Az oszlop- és sorindexből kiszámoljuk a parancs indexét

                if (opParam.settingsData.numOfSetCommands <= index) // Ha nem használt cellára kattintottak,
                    return; // akkor nem csinálunk semmit sem

                lock (timerQuery) // Letiltjuk a többi műveletet amíg nem végzünk itt
                {
                    DialogResult result;

                    if (opParam.settingsData.setCommands[index].numOfSections == 0) // Ha a parancs nem tartalmaz adatbájtokat,
                        result = DialogResult.Yes; // akkor nem kell megnyitni a parancshoz tartozó ablakot
                    else result = setDataWindows[index].ShowDialog(this); // Ha szükséges, akkor megnyitjuk a parancshoz tartozó ablakot

                    if (result == DialogResult.Yes) // Ha a felhasználó el akarja küldeni az adatot a műszernek,
                    { // akkor elküldjük
                        if (device.SendGeneral(comPort, (byte)pidGenerator.Next(1, 256), (byte)index, opParam, setDataWindows[index]))
                        { // Ha a beállító parancs sikeres volt,
                            setDataWindows[index].RefreshOldValues(); // akkor frissítjük a táblázatokat
                        }
                    }

                    dataReadBox.Text = device.lastReceivedData; // Az állapotsor frissítése
                    messageReadBox.Text = device.lastReceivedDataDecoded;
                }

                return;
            }
            catch (Exception ex) // Ha kivétel történt,
            {
                DateTime date = DateTime.Now; // akkor lementjük az időt,
                MessageBox.Show("An exception occured:" + Environment.NewLine + ex.Message); // majd megjelenítünk egy hibaüzenetet

                try
                {
                    fileWriter.WriteExceptionFile(date, "setCommandsTable_CellClick: " + ex.Message); // A hiba kiírása a kivétel-fájlba
                }
                catch (Exception exc) // Ha valamiért nem tudunk írni a kivétel-fájlba, akkor bezárjuk a programot
                {
                    MessageBox.Show("An exception occured during file writing:" + Environment.NewLine +
                                    exc.Message + Environment.NewLine + // akkor megjelenítünk egy hibaüzenetet
                                    "The application will terminate.");

                    this.Close(); // A program bezárása
                }

                return;
            }
        }
    }
}
