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
    /// <summary>Ez az osztály egy beállító parancs ablakát, illetve az ott található információkat tartalmazza</summary>
    /// <remarks>
    /// Az osztály nem rendelkezik teljes kivételkezeléssel, így ezt felsőbb szinten kell megoldani.
    /// Az egyes függvények leírásában található, hogy mely kivételeket kezelik le.
    /// </remarks>
    partial class SetData : Form
    {
        /// <summary>A beállító parancshoz tartozó adatok listája</summary>
        public List<SetSection> setSections;

        /// <summary>Keresőtáblázat a bitekhez annak megállapítására, hogy melyik adathoz tartoznak</summary>
        /// <remarks>
        /// A lista i-edik eleme azt mondja meg, hogy a setBitTable i-edig eleme hanyadik adathoz tartozik a setSections listában.
        /// Például ha sectionLookupTable[7] = 2, akkor az azt jelenti, hogy a setBitTable 7. indexű eleme a setSections[2] adathoz
        /// tartozik.
        /// A bit indexét a sor- és oszlopindexekből kell számolni, (sorindex * 3 + (oszlopindex / 3)) módon.
        /// Ha az érték -1, akkor az azt jelenti, hogy az adott táblázat-elemhez nem tartozik bit.
        /// </remarks>
        private List<int> sectionLookupTable;

        /// <summary>Keresőtáblázat a bitekhez annak megállapítására, hogy hanyadik bitről van szó</summary>
        /// <remarks>
        /// A lista i-edik eleme azt mondja meg, hogy a setBitTable i-edig eleme hanyadik bit az adott adatban
        /// Például ha bitLookupTable[3] = 1, akkor az azt jelenti, hogy a setBitTable 3. indexű eleme a
        /// setSections[sectionLookupTable[3]].bits[1] helyen található.
        /// A bit indexét a sor- és oszlopindexekből kell számolni, (sorindex * 3 + (oszlopindex / 3)) módon.
        /// Ha az érték -1, akkor az azt jelenti, hogy az adott táblázat-elemhez nem tartozik bit.
        /// </remarks>
        private List<int> bitLookupTable;

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

            /// <summary>Az adat legutolsó, felhasználó által megadott értéke</summary>
            /// <remarks>Ez az érték volt legutoljára beírva, és lementve a "Save and set" vagy "Save only" gombokkal.</remarks>
            public Object setValue;

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

            /// <summary>A bit legutolsó, felhasználó által megadott értéke</summary>
            /// <remarks>Ez az érték volt legutoljára beírva, és lementve a "Save and set" vagy "Save only" gombokkal.</remarks>
            public bool setValue;

            /// <summary>Igaz, ha a bit használatban van, különben hamis</summary>
            public bool isUsed;
        }

        /// <summary>Ez az osztály egy dátum/idő típusú cellát valósít meg</summary>
        public class CalendarCell : DataGridViewTextBoxCell
        {
            /// <summary>Konstruktor</summary>
            public CalendarCell()
                : base() // Az eredeti konsturktort használjuk, azzal a kitétellel,
            {
                this.Style.Format = "yyyy.MM.dd. HH\\:mm\\:ss"; // hogy a formátum dátum/idő
                return;
            }

            /// <summary>A cellában lévő vezérlő inicializálása</summary>
            /// <param name="rowIndex">A sorindex</param>
            /// <param name="initialFormattedValue">A cella kezdeti értéke</param>
            /// <param name="dataGridViewCellStyle">A cellában lévő vezérlő stílusa</param>
            public override void InitializeEditingControl(int rowIndex, object initialFormattedValue, DataGridViewCellStyle dataGridViewCellStyle)
            {
                // A vezérlő értéke a jelenlegi cellában lévő érték
                base.InitializeEditingControl(rowIndex, initialFormattedValue, dataGridViewCellStyle);
                CalendarEditingControl ctl = DataGridView.EditingControl as CalendarEditingControl;
                if (this.Value == null) // Ha a cellában nincs semmi,
                    ctl.Value = (DateTime)this.DefaultNewRowValue; // akkor az alapértelmezett értéket használjuk,
                else ctl.Value = (DateTime)this.Value; // különben pedig a jelenlegi értéket

                return;
            }

            /// <summary>A cellában lévő vezérlő típusának lekérdezése</summary>
            public override Type EditType
            {
                get
                {
                    return typeof(CalendarEditingControl);
                }
            }

            /// <summary>A cellában tárolt adat típusának lekérdezése</summary>
            public override Type ValueType
            {
                get
                {
                    return typeof(DateTime);
                }
            }

            /// <summary>A cellában tárolt adat alapértelmezett értéke</summary>
            public override object DefaultNewRowValue
            {
                get
                {
                    return DateTime.Now;
                }
            }
        }

        /// <summary>Egy dátum/idő vezérlőt megvalósító osztály</summary>
        class CalendarEditingControl : DateTimePicker, IDataGridViewEditingControl
        {
            /// <summary>A szülő objektum</summary>
            DataGridView dataGridView;

            /// <summary>Igaz, ha megváltozott a cella értéke</summary>
            private bool valueChanged = false;

            /// <summary>Sorindex</summary>
            int rowIndex;

            /// <summary>Konstruktor</summary>
            public CalendarEditingControl()
            {
                this.Format = DateTimePickerFormat.Custom; // Beállítjuk a dátum/idő formátumát
                this.CustomFormat = "yyyy.MM.dd. HH:mm:ss";
                return;
            }

            /// <summary>A cellában lévő érték lekérdezése és beállítása</summary>
            public object EditingControlFormattedValue
            {
                get // Az érték lekérdezése
                {
                    return this.Value.ToString("yyyy.MM.dd. HH\\:mm\\:ss");
                }
                set // Az érték beállítása
                {
                    if (value is String) // Ha a beállított érték sztring,
                    {
                        try
                        {
                            this.Value = DateTime.Parse((String)value); // akkor lemekérjük az értékét
                        }
                        catch // Ha az érték nem a megfelelő formátumú,
                        {
                            this.Value = DateTime.Now; // akkor az alapértelmezett értéket adjuk vissza
                        }
                    }
                    return;
                }
            }

            /// <summary>A cellában lévő érték lekérdezése</summary>
            /// <param name="context">Hibaleíró érték</param>
            /// <returns>A cellában lévő érték</returns>
            public object GetEditingControlFormattedValue(DataGridViewDataErrorContexts context)
            {
                return EditingControlFormattedValue;
            }

            /// <summary>A cella stílusának beállítása</summary>
            /// <param name="dataGridViewCellStyle">Az alkalmazandó stílus</param>
            public void ApplyCellStyleToEditingControl(DataGridViewCellStyle dataGridViewCellStyle)
            {
                this.Font = dataGridViewCellStyle.Font; // A betűtípus,
                this.CalendarForeColor = dataGridViewCellStyle.ForeColor; // valamint a színek beállítása
                this.CalendarMonthBackground = dataGridViewCellStyle.BackColor;
                return;
            }

            /// <summary>A sorindex lekérdezése/beállítása</summary>
            public int EditingControlRowIndex
            {
                get
                {
                    return rowIndex;
                }
                set
                {
                    rowIndex = value;
                    return;
                }
            }

            /// <summary>Ez a függvény dönti el, hogy a kapott billentyűparancsot végre kell-e hajtani, vagy sem</summary>
            /// <param name="key">A leütött billentyű</param>
            /// <param name="dataGridViewWantsInputKey">Igaz, ha a DataGridView vezérlő értelmezni tudja a parancsot</param>
            /// <returns>Igaz, ha az objektumnak végre kell hajtania a kapott parancsot</returns>
            public bool EditingControlWantsInputKey(Keys key, bool dataGridViewWantsInputKey)
            {
                // Let the DateTimePicker handle the keys listed. 
                switch (key & Keys.KeyCode) // Megvizsgáljuk a kapott billentyűparancsot,
                {
                    case Keys.Left: // és amit a DateTimePicker ismer,
                    case Keys.Up:
                    case Keys.Down:
                    case Keys.Right:
                    case Keys.Home:
                    case Keys.End:
                    case Keys.PageDown:
                    case Keys.PageUp:
                        return true; // azokat végrehajtjuk
                    default: // A többi parancsnál megvizsgáljuk,
                        return !dataGridViewWantsInputKey; // hogy a DataGriedView ismeri-e őket
                }
            }

            /// <summary>Ez a függvény előkészíti a cellát szerkesztésre</summary>
            /// <param name="selectAll">Igaz, ha a cella tartalmát ki kell jelölni</param>
            public void PrepareEditingControlForEdit(bool selectAll)
            {
                return; // Semmi előkészületet nem kell tenni
            }

            /// <summary>Annak beállítása, hogy a cella tartalmát újra kell-e pozícionálni, ha megváltozik annak tartalma</summary>
            public bool RepositionEditingControlOnValueChange
            {
                get
                {
                    return false; // A cella értékének változásakor sosincs szükség újrapozícionálásra
                }
            }

            /// <summary>A cella szülőobjektumának lekérdezése és beállítása</summary>
            public DataGridView EditingControlDataGridView
            {
                get
                {
                    return dataGridView;
                }
                set
                {
                    dataGridView = value;
                    return;
                }
            }

            /// <summary>Annak lekérdezése és beállítása, hogy a cella tartalma megváltozott-e</summary>
            public bool EditingControlValueChanged
            {
                get
                {
                    return valueChanged;
                }
                set
                {
                    valueChanged = value;
                    return;
                }
            }

            /// <summary>A cella tartalmának szerkesztésekor használt kurzor visszaadása</summary>
            public Cursor EditingPanelCursor
            {
                get
                {
                    return base.Cursor;
                }
            }

            /// <summary>A cella tartalmának megváltoztatásakor lefutó függvény</summary>
            /// <param name="eventargs">Az esemény paraméterei</param>
            protected override void OnValueChanged(EventArgs eventargs)
            {
                valueChanged = true; // Jelezzük a szülő objektum felé, hogy megváltozott a cella tartalma
                this.EditingControlDataGridView.NotifyCurrentCellDirty(true);
                base.OnValueChanged(eventargs);
                return;
            }
        }

        /// <summary>Konstruktor</summary>
        /// <param name="commandIndex">A beállító parancs sorszáma</param>
        /// <param name="opParamTmp">A használt paramétertáblázat</param>
        public SetData(byte commandIndex, OperatingParameters opParamTmp)
        {
            InitializeComponent(); // Az ablak inicializálása

            this.Name = opParamTmp.settingsData.setCommands[commandIndex].name; // Az ablak neve megegyezik a parancs nevével

            setSections = new List<SetSection>(opParamTmp.settingsData.numOfSetCommands); // Létrehozzuk a beállító parancsot, majd
            for (byte i = 0; i < opParamTmp.settingsData.setCommands[commandIndex].numOfSections; i++) // végigmegyünk a beállítható
            { // adatokon, és mindegyiket hozzáadjuk a beállítható adatok listájához,
                setSections.Add(new SetSection(opParamTmp.settingsData.setCommands[commandIndex].sections[i].numOfBits));

                setSections[i].name = opParamTmp.settingsData.setCommands[commandIndex].sections[i].name; // majd átmásoljuk a szükséges
                setSections[i].numOfBits = opParamTmp.settingsData.setCommands[commandIndex].sections[i].numOfBits; // információkat
                setSections[i].type = opParamTmp.settingsData.setCommands[commandIndex].sections[i].type;
                setSections[i].decimals = opParamTmp.settingsData.setCommands[commandIndex].sections[i].decimals;

                // Megvizsgáljuk az adat típusát, majd létrehozzuk
                if (setSections[i].type == typeof(DateTime)) // Az adat típusa dátum/idő
                    setSections[i].setValue = DateTime.Now;
                else if (setSections[i].type == typeof(sbyte)) // Az adat 8 bites előjeles egész
                    setSections[i].setValue = new sbyte();
                else if (setSections[i].type == typeof(short)) // Az adat 16 bites előjeles egész
                    setSections[i].setValue = new short();
                else if (setSections[i].type == typeof(int)) // Az adat 32 bites előjeles egész
                    setSections[i].setValue = new int();
                else if (setSections[i].type == typeof(byte)) // Az adat 8 bites előjel nélküli egész
                    setSections[i].setValue = new byte();
                else if (setSections[i].type == typeof(ushort)) // Az adat 16 bites előjel nélküli egész
                    setSections[i].setValue = new ushort();
                else // Az adat 32 bites előjel nélküli egész
                    setSections[i].setValue = new uint();

                for (byte j = 0; j < opParamTmp.settingsData.setCommands[commandIndex].sections[i].numOfBits; j++) // Végigmegyünk az
                { // aktuális adathoz tartozó biteken
                    setSections[i].bits.Add(new SetBit()); // és mindegyiket hozzáadjuk a bitek listájához, majd átmásoljuk a szükséges

                    setSections[i].bits[j].name = opParamTmp.settingsData.setCommands[commandIndex].sections[i].bits[j].name; // adatokat
                    setSections[i].bits[j].isUsed = opParamTmp.settingsData.setCommands[commandIndex].sections[i].bits[j].isUsed;

                    setSections[i].bits[j].setValue = false; // Létrehozzuk a tárolt adatot
                }
            }
            
            // Létrehozzuk a táblázatokat

            sectionLookupTable = new List<int>(); // A keresőtáblák létrehozása
            bitLookupTable = new List<int>();

            int dataRowNumber; // Az adatokat tartalmazó táblázat sorainak a száma
            dataRowNumber = setSections.Count / 3;
            if (setSections.Count % 3 != 0) dataRowNumber++; // Ha van "tört" sor, akkor 1-gyel növeljük a sorok számát

            if (dataRowNumber != 0) // Ha az adatokat megjelenítő táblázat tartalmaz sorokat,
                setDataTable.Rows.Add(dataRowNumber); // akkor inicializáljuk őket
            for (int i = 0; i < dataRowNumber; i++) // Végigmegyünk a sorokon
                setDataTable.Rows[i].Height = 20; // és beállítjuk azok magasságát

            int bitRowNumber = 0; // A biteket tartalmazó táblázat sorainak a száma
            for (byte i = 0; i < opParamTmp.settingsData.setCommands[commandIndex].numOfSections; i++) // Végigmegyünk a beállítható
            { // adatokon,
                byte bitNumber = 0; // valamint az aktuális adathoz tartozó biteken, és megszámoljuk a használt bitek számát
                for (byte j = 0; j < opParamTmp.settingsData.setCommands[commandIndex].sections[i].numOfBits; j++)
                {
                    if (setSections[i].bits[j].isUsed) // Ha az adott bit használatban van,
                    {
                        bitNumber++; // akkor növeljük a bitek számát
                    }
                }

                bitRowNumber += (bitNumber / 3); // Növeljük a sorok számát
                if ((bitNumber % 3) != 0) // Ha van "tört" sor, akkor
                    bitRowNumber++; // 1-gyel növeljük a sorok számát
            }

            if (bitRowNumber != 0) // Ha a biteket tartalmazó táblázat tartalmaz sorokat,
                setBitTable.Rows.Add(bitRowNumber); // akkor inicializáljuk őket
            for (int i = 0; i < bitRowNumber; i++) // Végigmegyünk a sorokon
                setBitTable.Rows[i].Height = 20; // és beállítjuk azok magasságát

            if (dataRowNumber != 0) // Ha vannak beállítandó adatok, akkor kiszámoljuk az
            { // az adatokat tartalmazó táblázat magasságának százalékos arányát a biteket tartalmazó táblázat magasságához képest
                int setDataTablePercent = (int)(((float)(dataRowNumber + 1) / ((float)(bitRowNumber + 1) + (float)(dataRowNumber + 1))) *
                    100.0F); // A plusz sorok a fejlécek miatt kellenek
                if (setDataTablePercent < 0) setDataTablePercent = 0; // A határok vizsgálata
                else if (100 < setDataTablePercent) setDataTablePercent = 100;

                tableLayoutPanel.RowStyles[0].Height = setDataTablePercent; // A magasságok arányában beállítjuk a táblázatok magasságát
                tableLayoutPanel.RowStyles[1].Height = (100 - setDataTablePercent);
            }

            // Feltöltjük a táblázatokat az adatokkal
            int num = 0; // Itt tároljuk a következő bit sorszámát
            for (byte i = 0; i < opParamTmp.settingsData.setCommands[commandIndex].numOfSections; i++) // Végigmegyünk a beállítható
            { // adatokon,
                setDataTable.Rows[i / 3].Cells[3 * (i % 3)].Value = setSections[i].name; // és feltöltjük adatokkal
                if (setSections[i].decimals == -2) // Ha az aktuális cella hexadecimális formátumú,
                { // akkor beállítjuk a bájtok számának megfelelő formátumot
                    if (setSections[i].type == typeof(byte)) // Az adat 8 bites előjel nélküli egész
                    {
                        setDataTable.Rows[i / 3].Cells[(3 * (i % 3)) + 1].Style.Format = "X2";
                        setDataTable.Rows[i / 3].Cells[(3 * (i % 3)) + 2].Style.Format = "X2";
                    }
                    else if (setSections[i].type == typeof(ushort)) // Az adat 16 bites előjel nélküli egész
                    {
                        setDataTable.Rows[i / 3].Cells[(3 * (i % 3)) + 1].Style.Format = "X4";
                        setDataTable.Rows[i / 3].Cells[(3 * (i % 3)) + 2].Style.Format = "X4";
                    }
                    else // Az adat 32 bites előjel nélküli egész
                    {
                        setDataTable.Rows[i / 3].Cells[(3 * (i % 3)) + 1].Style.Format = "X8";
                        setDataTable.Rows[i / 3].Cells[(3 * (i % 3)) + 2].Style.Format = "X8";
                    }

                    setDataTable.Rows[i / 3].Cells[(3 * (i % 3)) + 1].ReadOnly = true; // Itt nem lehet kézzel beállítani az értéket

                    setDataTable.Rows[i / 3].Cells[(3 * (i % 3)) + 1].Value = setSections[i].setValue; // Beírjuk az értékeket (kezdetben
                    setDataTable.Rows[i / 3].Cells[(3 * (i % 3)) + 2].Value = setSections[i].setValue; // mindkét érték is ugyanaz)
                }
                else if (setSections[i].decimals == -1) // Ha az aktuális cella dátum/idő formátumú,
                {
                    setDataTable.Rows[i / 3].Cells[(3 * (i % 3)) + 1] = new CalendarCell(); // akkor átalakítjuk a cellákat dátum/idő 
                    setDataTable.Rows[i / 3].Cells[(3 * (i % 3)) + 2] = new CalendarCell(); // típusúvá, majd beírjuk
                    setDataTable.Rows[i / 3].Cells[(3 * (i % 3)) + 1].Value = setSections[i].setValue; // az értékeket (kezdetben
                    setDataTable.Rows[i / 3].Cells[(3 * (i % 3)) + 2].Value = setSections[i].setValue; // mindkét érték is ugyanaz)
                }
                else // Az aktuális cella értéke egy "normál" (nem hexadecimális) szám (egész vagy tört)
                {
                    float setValueTmp; // A számok értéke (kezdetben mindkét érték is ugyanaz)
                    string format = "F"; // valamint formátuma

                    setValueTmp = Convert.ToSingle(setSections[i].setValue); // A számok értéke egészként

                    for (byte count = 0; count < setSections[i].decimals; count++) // Ahány tizedesjegyet tartalmaznak az értékek,
                        setValueTmp /= 10.0F; // annyiszor osztjuk le őket tízzel

                    // Kiegészítjük a formátum sztringet a tizedesjegyek számával
                    format += setSections[i].decimals.ToString();

                    setDataTable.Rows[i / 3].Cells[(3 * (i % 3)) + 1].Style.Format = format; // Beállítjuk a cellák formátumát,
                    setDataTable.Rows[i / 3].Cells[(3 * (i % 3)) + 2].Style.Format = format;

                    setDataTable.Rows[i / 3].Cells[(3 * (i % 3)) + 1].Value = setValueTmp; // majd beírjuk az értékeket (kezdetben
                    setDataTable.Rows[i / 3].Cells[(3 * (i % 3)) + 2].Value = setValueTmp; // mindkét érték is ugyanaz)
                }

                if ((num % 3) != 0) // Ha az előző sor tört sor volt a biteknél, akkor ugrunk a következő sorra
                {
                    for (byte tmp = 0; tmp < (3 - (num % 3)); tmp++) // Ahány elemet ugrunk
                    {
                        sectionLookupTable.Add(-1); // annyiszor -1-et írunk a keresőtáblába
                        bitLookupTable.Add(-1);
                    }

                    num += (3 - (num % 3)); // A következő elem sorszáma
                }

                for (byte j = 0; j < opParamTmp.settingsData.setCommands[commandIndex].sections[i].numOfBits; j++) // Végigmegyünk az
                { // adathoz tartozó biteken,
                    if (setSections[i].bits[j].isUsed) // és ha az adott bit használatban van,
                    {
                        setBitTable.Rows[num / 3].Cells[3 * (num % 3)].Value = setSections[i].bits[j].name; // akkor feltöltjük a táblázatot
                        ((DataGridViewCheckBoxCell)setBitTable.Rows[num / 3].Cells[(3 * (num % 3)) + 1]).Value = // a bithez tartozó
                            setSections[i].bits[j].setValue; // adatokkal (kezdetben mindkét érték is ugyanaz),
                        ((DataGridViewCheckBoxCell)setBitTable.Rows[num / 3].Cells[(3 * (num % 3)) + 2]).Value =
                            setSections[i].bits[j].setValue;

                        sectionLookupTable.Add(i); // valamint frissítjük a keresőtáblákat
                        bitLookupTable.Add(j);

                        num++; // A következő elem sorszáma
                    }
                }
            }

            return;
        }

        /// <summary>Ez a függvény frissíti a táblázatban lévő régi adatokat, és helyükre az új adatokat írja</summary>
        public void RefreshOldValues()
        {
            int num = 0; // Itt tároljuk a következő bit sorszámát
            for (byte i = 0; i < setSections.Count; i++) // Végigmegyünk a beállítható adatokon
            {
                if (setSections[i].decimals == -2) // Az aktuális cella hexadecimális formátumú
                {
                    setDataTable.Rows[i / 3].Cells[(3 * (i % 3)) + 2].Value = setSections[i].setValue; // Frissítjük a régi értéket
                }
                else if (setSections[i].decimals == -1) // Ha az aktuális cella dátum/idő formátumú,
                {
                    setDataTable.Rows[i / 3].Cells[(3 * (i % 3)) + 2].Value = setSections[i].setValue; // Frissítjük a régi értéket
                }
                else // Az aktuális cella értéke egy "normál" (nem hexadecimális) szám (egész vagy tört)
                {
                    float setValueTmp = Convert.ToSingle(setSections[i].setValue); // A szám értéke egészként

                    for (byte count = 0; count < setSections[i].decimals; count++) // Ahány tizedesjegyet tartalmaz az érték,
                        setValueTmp /= 10.0F; // annyiszor osztjuk le tízzel

                    setDataTable.Rows[i / 3].Cells[(3 * (i % 3)) + 2].Value = setValueTmp; // Frissítjük a régi értéket
                }

                if ((num % 3) != 0) // Ha az előző sor tört sor volt a biteknél, akkor ugrunk a következő sorra
                    num += (3 - (num % 3)); // A következő elem sorszáma

                for (byte j = 0; j < setSections[i].numOfBits; j++) // Végigmegyünk az adathoz tartozó biteken,
                {
                    if (setSections[i].bits[j].isUsed) // és ha az adott bit használatban van,
                    {
                        ((DataGridViewCheckBoxCell)setBitTable.Rows[num / 3].Cells[(3 * (num % 3)) + 2]).Value =
                            setSections[i].bits[j].setValue; // akkor frissítjük a régi értéket

                        num++; // A következő elem sorszáma
                    }
                }
            }

            return;
        }

        /// <summary>Ez a függvény leellenőrzi, hogy az adatok helyesek-e, majd lementi őket</summary>
        /// <param name="sender">Az eseményt küldő objektum</param>
        /// <param name="e">Az esemény paraméterei</param>
        /// <remarks>
        /// A "Save and set" és a "Save only" gombok is ezt a függvényt hívják meg.
        /// A függvény csak azokat a kivételeket kezeli, amelyek az adatok hibás megadásából következnek, a többi kivételt
        /// azonban továbbdobja, ezért a kivételkezelést felsőbb szinten kell megoldani.
        /// </remarks>
        private void saveOnlyButton_Click(object sender, EventArgs e)
        {
            int num = 0;
            for (byte i = 0; i < setSections.Count; i++) // Végigmegyünk a beállítható adatokon,
            {
                try
                {
                    // majd megvizsgáljuk az adat típusát, és e szerint lementjük az adatot
                    if (setSections[i].type == typeof(DateTime)) // Az adat típusa dátum/idő
                    {
                        setSections[i].setValue = (DateTime)(((CalendarCell)setDataTable.Rows[i / 3].Cells[(3 * (i % 3) + 1)]).Value);
                    }
                    else if (setSections[i].type == typeof(sbyte)) // Az adat 8 bites előjeles egész
                    { // Lementjük a táblázatban
                        float value = float.Parse(setDataTable.Rows[i / 3].Cells[(3 * (i % 3) + 1)].Value.ToString()); // tárolt értéket,

                        for (byte count = 0; count < setSections[i].decimals; count++) // majd ahány tizedesjegyet tartalmaz az érték,
                            value *= 10.0F; // annyiszor szorozzuk meg tízzel,

                        setSections[i].setValue = Convert.ToSByte(value); // végül pedig lementjük az egész értéket
                    }
                    else if (setSections[i].type == typeof(short)) // Az adat 16 bites előjeles egész
                    { // Lementjük a táblázatban
                        float value = float.Parse(setDataTable.Rows[i / 3].Cells[(3 * (i % 3) + 1)].Value.ToString()); // tárolt értéket,

                        for (byte count = 0; count < setSections[i].decimals; count++) // majd ahány tizedesjegyet tartalmaz az érték,
                            value *= 10.0F; // annyiszor szorozzuk meg tízzel,
                        
                        setSections[i].setValue = Convert.ToInt16(value); // végül pedig lementjük az egész értéket
                    }
                    else if (setSections[i].type == typeof(int)) // Az adat 32 bites előjeles egész
                    { // Lementjük a táblázatban
                        float value = float.Parse(setDataTable.Rows[i / 3].Cells[(3 * (i % 3) + 1)].Value.ToString()); // tárolt értéket,

                        for (byte count = 0; count < setSections[i].decimals; count++) // majd ahány tizedesjegyet tartalmaz az érték,
                            value *= 10.0F; // annyiszor szorozzuk meg tízzel,
                        
                        setSections[i].setValue = Convert.ToInt32(value); // végül pedig lementjük az egész értéket
                    }
                    else if (setSections[i].type == typeof(byte)) // Az adat 8 bites előjel nélküli egész
                    {
                        if (setSections[i].decimals == -2) // Ha az érték hexadecimális formátumban van,
                        { // akkor egyszerűen csak lementjük a táblázatban tárolt értéket
                            setSections[i].setValue = Convert.ToByte(setDataTable.Rows[i / 3].Cells[(3 * (i % 3) + 1)].Value.ToString(), 16);
                        }
                        else // Ha a szám "normál" (nem hexadecimális) formátumban van, akkor a lementjük a táblázatban
                        {
                            float value = float.Parse(setDataTable.Rows[i / 3].Cells[(3 * (i % 3) + 1)].Value.ToString()); // tárolt értéket,

                            for (byte count = 0; count < setSections[i].decimals; count++) // majd ahány tizedesjegyet tartalmaz az érték,
                                value *= 10.0F; // annyiszor szorozzuk meg tízzel,

                            setSections[i].setValue = Convert.ToByte(value); // végül pedig lementjük az egész értéket
                        }
                    }
                    else if (setSections[i].type == typeof(ushort)) // Az adat 16 bites előjel nélküli egész
                    {
                        if (setSections[i].decimals == -2) // Ha az érték hexadecimális formátumban van,
                        { // akkor egyszerűen csak lementjük a táblázatban tárolt értéket
                            setSections[i].setValue = Convert.ToUInt16(setDataTable.Rows[i / 3].Cells[(3 * (i % 3) + 1)].Value.ToString(), 16);
                        }
                        else // Ha a szám "normál" (nem hexadecimális) formátumban van, akkor a lementjük a táblázatban
                        {
                            float value = float.Parse(setDataTable.Rows[i / 3].Cells[(3 * (i % 3) + 1)].Value.ToString()); // tárolt értéket,

                            for (byte count = 0; count < setSections[i].decimals; count++) // majd ahány tizedesjegyet tartalmaz az érték,
                                value *= 10.0F; // annyiszor szorozzuk meg tízzel,

                            setSections[i].setValue = Convert.ToUInt16(value); // végül pedig lementjük az egész értéket
                        }
                    }
                    else // Az adat 32 bites előjel nélküli egész
                    {
                        if (setSections[i].decimals == -2) // Ha az érték hexadecimális formátumban van,
                        { // akkor egyszerűen csak lementjük a táblázatban tárolt értéket
                            setSections[i].setValue = Convert.ToUInt32(setDataTable.Rows[i / 3].Cells[(3 * (i % 3) + 1)].Value.ToString(), 16);
                        }
                        else // Ha a szám "normál" (nem hexadecimális) formátumban van, akkor a lementjük a táblázatban
                        {
                            float value = float.Parse(setDataTable.Rows[i / 3].Cells[(3 * (i % 3) + 1)].Value.ToString()); // tárolt értéket,

                            for (byte count = 0; count < setSections[i].decimals; count++) // majd ahány tizedesjegyet tartalmaz az érték,
                                value *= 10.0F; // annyiszor szorozzuk meg tízzel,

                            setSections[i].setValue = Convert.ToUInt32(value); // végül pedig lementjük az egész értéket
                        }
                    }

                    if ((num % 3) != 0) // Ha az előző sor tört sor volt a biteknél, akkor ugrunk a következő sorra
                        num += (3 - (num % 3));

                    for (byte j = 0; j < setSections[i].numOfBits; j++) // Végigmegyünk az adathoz tartozó biteken,
                    {
                        if (setSections[i].bits[j].isUsed) // és ha az adott bit használatban van,
                        {
                            setSections[i].bits[j].setValue = // akkor lementjük az értékét
                                (bool)(((DataGridViewCheckBoxCell)setBitTable.Rows[num / 3].Cells[(3 * (num % 3) + 1)]).Value);

                            num++; // A következő bit sorszáma
                        }
                    }
                }
                catch (FormatException) // Ha nem megfelelő a felhasználó által megadott adat formátuma,
                {
                    string message = "An exception occurred:" + Environment.NewLine + // akkor létrehozunk,
                        "The value in row " + ((i / 3) + 1).ToString() + ", column " + (3 * (i % 3) + 2).ToString() +
                        " has incorrect format." + Environment.NewLine + "Please correct the value in order to continue!";

                    MessageBox.Show(message); // majd megjelenítünk egy hibaüzenetet

                    this.DialogResult = DialogResult.None; // Mivel az adatok nem megfelelőek, ezért nem zárjuk még be az ablakot
                    return;
                }
                catch (OverflowException) // Ha a felhasználó által megadott adat nincs a megfelelő határokon belül,
                {
                    string message = "An exception occurred:" + Environment.NewLine + // akkor létrehozunk egy hibaüzenetet
                        "The value in row " + ((i / 3) + 1).ToString() + ", column " + (3 * (i % 3) + 2).ToString() + " is out of range." +
                        Environment.NewLine + "The value should no less, than ";

                    // A hibaüzenetben megjelenítjük a határokat
                    if (setSections[i].type == typeof(sbyte)) // Az adat 8 bites előjeles egész
                    {
                        message += (((float)(sbyte.MinValue)) / (Math.Pow(10, setSections[i].decimals))).ToString() +
                            " and no greater, than " + (((float)(sbyte.MaxValue)) / (Math.Pow(10, setSections[i].decimals))).ToString();
                    }
                    else if (setSections[i].type == typeof(short)) // Az adat 16 bites előjeles egész
                    {
                        message += (((float)(short.MinValue)) / (Math.Pow(10, setSections[i].decimals))).ToString() +
                            " and no greater, than " + (((float)(short.MaxValue)) / (Math.Pow(10, setSections[i].decimals))).ToString();
                    }
                    else if (setSections[i].type == typeof(int)) // Az adat 32 bites előjeles egész
                    {
                        message += (((float)(int.MinValue)) / (Math.Pow(10, setSections[i].decimals))).ToString() +
                            " and no greater, than " + (((float)(int.MaxValue)) / (Math.Pow(10, setSections[i].decimals))).ToString();
                    }
                    else if (setSections[i].type == typeof(byte)) // Az adat 8 bites előjel nélküli egész
                    {
                        message += (((float)(byte.MinValue)) / (Math.Pow(10, setSections[i].decimals))).ToString() +
                            " and no greater, than " + (((float)(byte.MaxValue)) / (Math.Pow(10, setSections[i].decimals))).ToString();
                    }
                    else if (setSections[i].type == typeof(ushort)) // Az adat 16 bites előjel nélküli egész
                    {
                        message += (((float)(ushort.MinValue)) / (Math.Pow(10, setSections[i].decimals))).ToString() +
                            " and no greater, than " + (((float)(ushort.MaxValue)) / (Math.Pow(10, setSections[i].decimals))).ToString();
                    }
                    else // Az adat 32 bites előjel nélküli egész
                    {
                        message += (((float)(uint.MinValue)) / (Math.Pow(10, setSections[i].decimals))).ToString() +
                            " and no greater, than " + (((float)(uint.MaxValue)) / (Math.Pow(10, setSections[i].decimals))).ToString();
                    }

                    message += Environment.NewLine + "Please correct the value in order to continue!"; // Kiegészítjük,

                    MessageBox.Show(message); // majd megjelenítünk a hibaüzenetet

                    this.DialogResult = DialogResult.None; // Mivel az adatok nem megfelelőek, ezért nem zárjuk még be az ablakot
                    return;
                }
                // A többi kivételt továbbdobjuk
            }

            return;
        }

        /// <summary>Akkor fut le, ha kijelöltek egy cellát abban a táblázatban, ahol a beállítható adatok találhatóak</summary>
        /// <param name="sender">Az eseményt küldő objektum</param>
        /// <param name="e">Az esemény paraméterei</param>
        /// <remarks>A függvény nem rendelkezik teljes kivételkezeléssel, így ezt felsőbb szinten kell megoldani.</remarks>
        private void setDataTable_SelectionChanged(object sender, EventArgs e)
        {
            if (setDataTable.SelectedCells.Count > 0) // Ha kijelöltek legalább egy cellát,
                setDataTable.SelectedCells[0].Selected = false; // akkor megszüntetjük a kijelölést
            return;
        }

        /// <summary>Akkor fut le, ha kijelöltek egy cellát abban a táblázatban, ahol a beállítható bitek találhatóak</summary>
        /// <param name="sender">Az eseményt küldő objektum</param>
        /// <param name="e">Az esemény paraméterei</param>
        /// <remarks>A függvény nem rendelkezik teljes kivételkezeléssel, így ezt felsőbb szinten kell megoldani.</remarks>
        private void setBitTable_SelectionChanged(object sender, EventArgs e)
        {
            if (setBitTable.SelectedCells.Count > 0) // Ha kijelöltek legalább egy cellát,
                setBitTable.SelectedCells[0].Selected = false; // akkor megszüntetjük a kijelölést
            return;
        }

        /// <summary>Akkor fut le, ha rákattintottak egy cellára abban a táblázatban, ahol a beállítható bitek találhatóak</summary>
        /// <param name="sender">Az eseményt küldő objektum</param>
        /// <param name="e">Az esemény paraméterei</param>
        /// <remarks>A függvény nem rendelkezik teljes kivételkezeléssel, így ezt felsőbb szinten kell megoldani.</remarks>
        private void setBitTable_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex != 1 && e.ColumnIndex != 4 && e.ColumnIndex != 7) // Ha nem a beállítható oszlopokba kattintottak,
                return; // akkor nem teszünk semmit sem

            int index = e.RowIndex * 3 + (e.ColumnIndex / 3); // Az oszlop- és sorindexből kiszámoljuk a bit indexét

            if (bitLookupTable.Count <= index || bitLookupTable[index] == -1) // Ha egy nem használt jelölőnégyzet cellájára kattintottak,
            { // akkor a jelölőnégyzetet nem
                ((DataGridViewCheckBoxCell)setBitTable.Rows[e.RowIndex].Cells[e.ColumnIndex]).Value = false; // engedjük bepipálni
            }
            else // Egy használatban lévő jelölőnégyzet cellájára kattintottak,
            {
                DataGridViewCheckBoxCell cell = (DataGridViewCheckBoxCell)setBitTable.Rows[e.RowIndex].Cells[e.ColumnIndex];

                if ((bool)(cell.Value)) // Ha a jelölőnégyzet be van pipálva,
                {
                    cell.Value = false; // akkor töröljük a pipát, majd megvizsgáljuk az adat típusát

                    int bitIndex = bitLookupTable[index]; // Megnézzük, hogy hanyadik bitet módosítottuk

                    if (setSections[sectionLookupTable[index]].type == typeof(byte)) // Az adat 8 bites előjel nélküli egész
                    { // Kiolvassuk a számot az adatok táblázatából,
                        byte num = Convert.ToByte(setDataTable.Rows[sectionLookupTable[index] / 3].
                            Cells[(3 * (sectionLookupTable[index] % 3) + 1)].Value.ToString(), 16);
                        num &= (byte)(~(1 << (7 - bitIndex))); // majd frissítjük az értéket,
                        setDataTable.Rows[sectionLookupTable[index] / 3].Cells[(3 * (sectionLookupTable[index] % 3) + 1)].Value =
                            num.ToString("X2"); // és visszaírjuk a táblázatba
                    }
                    else if (setSections[sectionLookupTable[index]].type == typeof(ushort)) // Az adat 16 bites előjel nélküli egész
                    { // Kiolvassuk a számot az adatok táblázatából,
                        ushort num = Convert.ToUInt16(setDataTable.Rows[sectionLookupTable[index] / 3].
                            Cells[(3 * (sectionLookupTable[index] % 3) + 1)].Value.ToString(), 16);
                        num &= (ushort)(~(1 << (15 - bitIndex))); // majd frissítjük az értéket,
                        setDataTable.Rows[sectionLookupTable[index] / 3].Cells[(3 * (sectionLookupTable[index] % 3) + 1)].Value =
                            num.ToString("X4"); // és visszaírjuk a táblázatba
                    }
                    else // Az adat 32 bites előjel nélküli egész
                    { // Kiolvassuk a számot az adatok táblázatából,
                        uint num = Convert.ToUInt32(setDataTable.Rows[sectionLookupTable[index] / 3].
                            Cells[(3 * (sectionLookupTable[index] % 3) + 1)].Value.ToString(), 16);
                        num &= (uint)(~(1 << (31 - bitIndex))); // majd frissítjük az értéket,
                        setDataTable.Rows[sectionLookupTable[index] / 3].Cells[(3 * (sectionLookupTable[index] % 3) + 1)].Value =
                            num.ToString("X8"); // és visszaírjuk a táblázatba
                    }
                }
                else // Ha a jelölőnégyzet nincs bepipálva,
                {
                    cell.Value = true; // akkor bepipáljuk, majd megvizsgáljuk az adat típusát

                    int bitIndex = bitLookupTable[index]; // Megnézzük, hogy hanyadik bitet módosítottuk

                    if (setSections[sectionLookupTable[index]].type == typeof(byte)) // Az adat 8 bites előjel nélküli egész
                    { // Kiolvassuk a számot az adatok táblázatából,
                        byte num = Convert.ToByte(setDataTable.Rows[sectionLookupTable[index] / 3].
                            Cells[(3 * (sectionLookupTable[index] % 3) + 1)].Value.ToString(), 16);
                        num |= (byte)(1 << (7 - bitIndex)); // majd frissítjük az értéket,
                        setDataTable.Rows[sectionLookupTable[index] / 3].Cells[(3 * (sectionLookupTable[index] % 3) + 1)].Value =
                            num.ToString("X2"); // és visszaírjuk a táblázatba
                    }
                    else if (setSections[sectionLookupTable[index]].type == typeof(ushort)) // Az adat 16 bites előjel nélküli egész
                    { // Kiolvassuk a számot az adatok táblázatából,
                        ushort num = Convert.ToUInt16(setDataTable.Rows[sectionLookupTable[index] / 3].
                            Cells[(3 * (sectionLookupTable[index] % 3) + 1)].Value.ToString(), 16);
                        num |= (ushort)(1 << (15 - bitIndex)); // majd frissítjük az értéket,
                        setDataTable.Rows[sectionLookupTable[index] / 3].Cells[(3 * (sectionLookupTable[index] % 3) + 1)].Value =
                            num.ToString("X4"); // és visszaírjuk a táblázatba
                    }
                    else // Az adat 32 bites előjel nélküli egész
                    { // Kiolvassuk a számot az adatok táblázatából,
                        uint num = Convert.ToUInt32(setDataTable.Rows[sectionLookupTable[index] / 3].
                            Cells[(3 * (sectionLookupTable[index] % 3) + 1)].Value.ToString(), 16);
                        num |= (uint)(1 << (31 - bitIndex)); // majd frissítjük az értéket,
                        setDataTable.Rows[sectionLookupTable[index] / 3].Cells[(3 * (sectionLookupTable[index] % 3) + 1)].Value =
                            num.ToString("X8"); // és visszaírjuk a táblázatba
                    }
                }
            }

            return;
        }
    }
}