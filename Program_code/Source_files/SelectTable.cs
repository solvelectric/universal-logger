using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace universal_logger
{
    /// <summary>Ez az osztály tartalmazza az inicializációs ablakot, azaz azt az ablakot, ahol kiválasztható a paramétertáblázat</summary>
    /// <remarks>
    /// Az osztály végigszkenneli a megadott könyvtárat, majd egy listában megjeleníti az elérhető táblázatokat, amelyből
    /// a felhasználó kiválaszthatja a megfelelőt, majd elindíthatja a főprogramot.
    /// Az osztály nem rendelkezik kivételkezeléssel, így ezt felsőbb szinten kell megoldani.
    /// </remarks>
    partial class SelectTable : Form
    {
        /// <summary>A felhasználó által kiválasztott paramétertáblázat elérési útvonala</summary>
        public string selectedTable;

        /// <summary>Igaz, ha a tábla automatikusan lett kiválasztva</summary>
        public bool autoSelected;

        /// <summary>A kiválasztható paramétertáblázatok listája</summary>
        private string[] parameterFiles;

        /// <summary>Az alapértelmezett táblázat indexe a parameterFiles tömbben. -1, ha nincs alapértelmezett táblázat</summary>
        private int defaultTableIndex;

        /// <summary>Időzítő az ablak bezárására, és a táblázat automatikus kiválasztására</summary>
        private Timer timer;

        /// <summary>Ennyi idő van még hátra az ablak automatikus bezárásáig, és a táblázat kiválasztásáig</summary>
        private int remainingSeconds;

        /// <summary>Alapértelmezett konstruktor</summary>
        /// <remarks>Azt a konstruktort kell használni e helyett, amelyikben megadható a program elérési útvonala.</remarks>
        public SelectTable()
        {
            InitializeComponent(); // Az ablak inicializálása
        }

        /// <summary>Konstruktor</summary>
        /// <param name="path">A program elérési útvonala</param>
        /// <remarks>
        /// A függvény nem rendelkezik kivételkezeléssel, így a hívó függvénynek kell ezt megoldania.
        /// </remarks>
        public SelectTable(string path)
        {
            InitializeComponent(); // Az ablak inicializálása
            
            parameterFiles = Directory.GetFiles(path, "*.xlsx"); // Lekérdezzük a könyvtárból az xlsx kiterjesztésű fájlokat
            Array.Sort(parameterFiles); // Rendezzük a tömböt

            defaultTableIndex = -1; // Még nem találtuk meg az alapértelmezett táblázatot
            autoSelected = false;
            for (int i = 0; i < parameterFiles.Length; i++) // Végigmegyünk az összes fájlon,
            { // levágjuk az elérési útvonalat a fájlnévről,
                string tableName = parameterFiles[i].Substring(path.Length + 1);
                tableList.Items.Add(tableName); // majd betesszük a listába

                if (tableName == "default_table.xlsx") // Ha a táblázat az alapértelmezett táblázat,
                    defaultTableIndex = i; // akkor lementjük
            }

            if (defaultTableIndex != -1) // Ha van alapértelmezett táblázat,
            {
                timer = new Timer(); // akkor létrehozzuk a bezáráshoz szükséges időzítőt
                timer.Interval = 1000; // 1 másodperces időtúllépéssel
                timer.Tick += timerEvent;
                timer.Start();
                remainingSeconds = 30; // Fél perc múlva zárjuk be az ablakot
                
                // Frissítjük a feliratot az ablakban
                remainingTimeLabel.Text = "Selecting default_table.xlsx in " + remainingSeconds.ToString() + " seconds.";
            }

            return;
        }

        /// <summary>Ez a függvény fut le, ha a felhasználó a "Start" feliratú gombra kattint</summary>
        /// <param name="sender">Az eseményt küldő objektum</param>
        /// <param name="e">Az esemény paraméterei</param>
        /// <remarks>
        /// A függvény nem rendelkezik kivételkezeléssel, így a hívó függvénynek kell ezt megoldania.
        /// </remarks>
        private void StartButton_Click(object sender, EventArgs e)
        {
            if (tableList.SelectedItem == null) // Ha a felhasználó nem választott ki semmit
            { // akkor megjelenítünk egy hibaüzenetet,
                MessageBox.Show("You have to choose a parameter-table from the list!" + Environment.NewLine +
                                "If the list is empty, please contact with the system administrator!");
                return; // és még nem zárjuk be az ablakot
            }

            if (defaultTableIndex != -1) // Ha van alapértelmezett táblázat,
            {
                timer.Stop(); // akkor leállítjuk az időzítőt
            }
            
            // Ha a felhasználó kiválasztotta a paramétertáblázatot,
            selectedTable = parameterFiles[tableList.SelectedIndex]; // akkor beállítjuk a kiválasztott tábla elérési útvonalát,
            this.DialogResult = DialogResult.OK; // majd elindítjuk a főablakot, jelezve, hogy ki lett választva a táblázat

            this.Close(); // Bezárjuk az inicializációs ablakot
            return;
        }

        /// <summary>Az időzítő jelzésekor lefutó függvény</summary>
        /// <param name="sender">Az eseményt küldő objektum</param>
        /// <param name="e">Az esemény paraméterei</param>
        private void timerEvent(Object sender, EventArgs e)
        {
            remainingSeconds--; // Csökkentjük a hátralévő időt

            if (remainingSeconds <= 0) // Ha letelt az idő,
            {
                timer.Stop(); // akkor leállítjuk az időzítőt,
                selectedTable = parameterFiles[defaultTableIndex]; // beállítjuk az alapértelmezett táblázatot,
                autoSelected = true; // jelezzük, hogy automatikus táblázat-kiválasztás történt,
                this.DialogResult = DialogResult.OK; // majd bezárjuk az ablakot
                this.Close();
            }
            else // Ha még nem telt le az idő, akkor frissítjük az ablakban lévő feliratot
                remainingTimeLabel.Text = "Selecting default_table.xlsx in " + remainingSeconds.ToString() + " seconds.";

            return;
        }
    }
}
