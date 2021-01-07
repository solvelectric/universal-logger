using System;
using System.Collections.Generic;
using System.IO; // A fájlkezeléshez szükséges
using System.Linq;
using System.Net; // Az e-mail küldés,
using System.Net.Mail; // és az ezzel kapcsolatos műveletek miatt kellenek
using System.Text;

namespace universal_logger
{
    /// <summary>Ez az osztály kezeli a lemezműveleteket, valamint az e-mail-ek küldését.</summary>
    /// <remarks>Az osztály nem rendelkezik kivételkezeléssel, így ezt felsőbb szinten kell megoldani.</remarks>
    class TextFileWriter
    {
        /// <summary>A program elérési útvonala</summary>
        private string applicationPath;

        /// <summary>A kivétel-fájl elérési útvonala</summary>
        private string exceptionFilePath;

        /// <summary>A logfájl elérési útvonala</summary>
        private string logFilePath;

        /// <summary>A logfájlok könyvtárának az elérési útvonala</summary>
        private string logDirectoryPath;

        /// <summary>Az utolsó fájl létrehozásának dátuma/ideje</summary>
        private DateTime lastCreate;

        /// <summary>A logfájl fejléce</summary>
        private string logFileHeader;

        /// <summary>Konstruktor a program elérési útvonalával</summary>
        /// <param name="appPath">A program elérési útvonala</param>
        /// <remarks>
        /// Ezt a konstruktort kell használni az alapértelmezett konstruktor helyett.
        /// A konstruktor nem rendelkezik kivételkezeléssel, így a hívó függvénynek kell ezt megoldania.
        /// </remarks>
        public TextFileWriter(string appPath)
        {
            applicationPath = appPath; // A program elérési útvonala
            exceptionFilePath = appPath + Path.DirectorySeparatorChar + "exceptions.txt"; // A kivétel-fájl elérési útvonalának inicializálása
            return;
        }

        /// <summary>A logfájl paramétereinek a beállítása</summary>
        /// <param name="logDirPath">A logfájlok könyvtárának az elérési útvonala</param>
        /// <param name="logHeader">A logfájl fejléce</param>
        public void setLogFileParameters(string logDirPath, string logHeader)
        {
            logDirectoryPath = logDirPath; // A logfájlok könyvtárának a neve
            logFileHeader = logHeader; // A logfájl fejléce

            CreateFile(); // A könyvtárszerkezet, és a logfájlok inicializálása

            return;
        }

        /// <summary>Ez a függvény ha kell, akkor létrehozza a könyvtárstruktúrát, valamint a logfájlt</summary>
        /// <remarks>
        /// Mivel a függvény nem írja felül a már létező fájlokat, így biztonságosan meghívható.
        /// A függvény nem rendelkezik kivételkezeléssel, így a hívó függvénynek kell ezt megoldania.
        /// </remarks>
        private void CreateFile()
        {
            DateTime nowTmp = DateTime.Now; // Lementjük az aktuális dátumot

            string yearString, monthString, dayString; // A fájlnévhez szükséges változók

            yearString = nowTmp.Year.ToString(); // Lementjük az évszámot
            if (nowTmp.Month < 10) monthString = "0" + nowTmp.Month.ToString(); // Lementjük a hónap számát, ha kell, elétesszük
            else monthString = nowTmp.Month.ToString(); // a '0'-t is
            if (nowTmp.Day < 10) dayString = "0" + nowTmp.Day.ToString(); // Lementjük a nap számát, ha kell, elétesszük a '0'-t is
            else dayString = nowTmp.Day.ToString();

            // A logfájl elérési útvonala
            logFilePath = logDirectoryPath + Path.DirectorySeparatorChar + "L" + yearString + monthString + dayString + ".csv";

            DirectoryInfo logDirInfo = new DirectoryInfo(logDirectoryPath); // Ellenőrizzük, hogy létezik-e már a
            if (!logDirInfo.Exists) logDirInfo.Create(); // könyvtár, és ha kell, akkor létrehozzuk

            FileInfo logFileInfo = new FileInfo(logFilePath); // Lekérjük a fájlinformációkat, és ha még nem létezik a fájl, akkor
            if (logFileInfo.Exists == false) File.WriteAllText(logFilePath, logFileHeader); // létrehozzuk, és beleírjuk az első sort

            // Aktualizáljuk az utolsó fájl létrehozásának idejét (ha nem kell létrehozni semmit sem,
            lastCreate = nowTmp; // akkor is aktuálizálni kell az értéket)
        }

        /// <summary>Ez a függvény ha kell, akkor létrehozza a könyvtárstruktúrát, valamint a logfájlt a megadott dátum/idő alapján</summary>
        /// <param name="date">A használandó dátum/idő</param>
        /// <remarks>
        /// Mivel a függvény nem írja felül a már létező fájlokat, így biztonságosan meghívható.
        /// A függvény nem rendelkezik kivételkezeléssel, így a hívó függvénynek kell ezt megoldania.
        /// </remarks>
        public void CreateFile(DateTime date)
        {
            string yearString, monthString, dayString; // A fájlnévhez szükséges változók

            yearString = date.Year.ToString(); // Lementjük az évszámot
            if (date.Month < 10) monthString = "0" + date.Month.ToString(); // Lementjük a hónap számát, ha kell, elétesszük
            else monthString = date.Month.ToString(); // a '0'-t is
            if (date.Day < 10) dayString = "0" + date.Day.ToString(); // Lementjük a nap számát, ha kell, elétesszük a '0'-t is
            else dayString = date.Day.ToString();

            // A logfájl elérési útvonala
            logFilePath = logDirectoryPath + Path.DirectorySeparatorChar + "L" + yearString + monthString + dayString + ".csv";

            DirectoryInfo logDirInfo = new DirectoryInfo(logDirectoryPath); // Ellenőrizzük, hogy létezik-e már a
            if (!logDirInfo.Exists) logDirInfo.Create(); // könyvtár, és ha kell, akkor létrehozzuk

            FileInfo logFileInfo = new FileInfo(logFilePath); // Lekérjük a fájlinformációkat, és ha még nem létezik a fájl, akkor
            if (logFileInfo.Exists == false) File.WriteAllText(logFilePath, logFileHeader); // létrehozzuk, és beleírjuk az első sort

            // Aktualizáljuk az utolsó fájl létrehozásának idejét (ha nem kell létrehozni semmit sem,
            lastCreate = date; // akkor is aktuálizálni kell az értéket)
        }

        /// <summary>Ez a függvény kiírja a kapott szöveget a kivétel-fájlba</summary>
        /// <param name="date">A sor kiírásának dátuma/ideje</param>
        /// <param name="str">A kiírandó sor</param>
        /// <remarks>A függvény nem rendelkezik kivételkezeléssel, így a hívó függvénynek kell ezt megoldania.</remarks>
        public void WriteExceptionFile(DateTime date, string str)
        {
            string line = date.ToString() + "\t" + str + Environment.NewLine; // A sor létrehozása
            File.AppendAllText(exceptionFilePath, line); // A sor kiírása a fájlba, szükség esetén a fájl létrehozása
        }

        /// <summary>Ez a függvény kiírja a kapott szöveget a logfájlba</summary>
        /// <param name="date">A sor kiírásának dátuma/ideje</param>
        /// <param name="str">A kiírandó sor</param>
        /// <remarks>A függvény nem rendelkezik kivételkezeléssel, így a hívó függvénynek kell ezt megoldania.</remarks>
        public void WriteLogFile(DateTime date, string str)
        {
            if (date.Date != lastCreate.Date) // Ha megváltozott a dátum,
                CreateFile(date); // akkor létrehozzuk az új fájlt

            File.AppendAllText(logFilePath, str); // A sor kiírása a fájlba, szükség esetén a fájl létrehozása
        }

        /// <summary>Ez a függvény küldi el a logfájlt e-mailben</summary>
        /// <param name="date">A küldés dátuma</param>
        /// <param name="appName">A program neve (verziószámmal)</param>
        /// <param name="sender">A küldő</param>
        /// <param name="addresses">A címzettek</param>
        /// <returns>A visszatérési érték igaz, ha az e-mail elküldése sikeres volt, hamis, ha nem sikerült a küldés</returns>
        /// <remarks>
        /// Az elküldött fájl a küldés dátuma előtti nap logfájlja lesz
        /// A függvény nem rendelkezik kivételkezeléssel, így a hívó függvénynek kell ezt megoldania.
        /// </remarks>
        public bool SendDailyMail(DateTime date, string appName, string sender, MailAddressCollection addresses)
        {
            string subjectTmp, messageTmp;
            List<Attachment> attachmentsTmp = new List<Attachment>();

            subjectTmp = "Daily e-mail by " + sender + ", " + date.ToShortDateString(); // Az üzenet tárgya

            messageTmp = // Az üzenet szövege
                "Dear Sir/Madam," + Environment.NewLine +
                Environment.NewLine +
                "this e-mail is automatically sended to you by " + sender + ", please don't reply to it!" + Environment.NewLine +
                "The message was generated by the " + appName + " software." + Environment.NewLine +
                Environment.NewLine +
                "System messages:" + Environment.NewLine;

            string yearString, monthString, dayString; // A fájlnévhez szükséges változók

            yearString = date.AddDays(-1).Year.ToString(); // Lementjük az előző napi évszámot
            if (date.AddDays(-1).Month < 10) monthString = "0" + date.AddDays(-1).Month.ToString(); // Lementjük az előző naphoz tartozó
            else monthString = date.AddDays(-1).Month.ToString(); // hónap számát, és ha kell, elétesszük a '0'-t is
            if (date.AddDays(-1).Day < 10) dayString = "0" + date.AddDays(-1).Day.ToString(); // Lementjük az előző nap számát, és ha kell,
            else dayString = date.AddDays(-1).Day.ToString(); // elétesszük a '0'-t is

            string logFilePathTmp; // Az ideiglenes fájl neve

            // Az előző napi logfájl elérési útvonala
            logFilePathTmp = logDirectoryPath + Path.DirectorySeparatorChar + "L" + yearString + monthString + dayString + ".csv";

            FileInfo logFileInfo = new FileInfo(logFilePathTmp); // Lekérjük a fájlinformációkat a fájlról

            if (logFileInfo.Exists == true && logFileInfo.Length <= 5242880) // Ha létezik a fájl, és kisebb,
            {
                attachmentsTmp.Add(new Attachment(logFilePathTmp)); // mint 5Mb, akkor betesszük a mellékletek közé
                messageTmp += "The logfile is attached!" + Environment.NewLine;
            }
            else if (logFileInfo.Exists == false) // Nem létezik a fájl
                messageTmp += "The logfile does not exist!" + Environment.NewLine;
            else // A fájl túl nagy a küldéshez
                messageTmp += "The logfile is too big to attach! File size: " +
                    logFileInfo.Length.ToString() + " bytes" + Environment.NewLine;

            messageTmp += // Az üzenet szövegének a vége
                Environment.NewLine +
                "If you don't want to receive these e-mails, please contact with the system administrator." + Environment.NewLine;

            return SendMail(subjectTmp, messageTmp, attachmentsTmp, addresses); // Elküldjük a levelet
        }

        /// <summary>Ez a függvény e-mail küld a megadott kivételről</summary>
        /// <param name="date">A küldés dátuma</param>
        /// <param name="appName">A program neve (verziószámmal)</param>
        /// <param name="sender">A küldő</param>
        /// <param name="addresses">A címzettek</param>
        /// <returns>A visszatérési érték igaz, ha az e-mail elküldése sikeres volt, hamis, ha nem sikerült a küldés</returns>
        /// <remarks>
        /// Az elküldött fájl a küldés dátuma előtti nap logfájlja lesz
        /// A függvény nem rendelkezik kivételkezeléssel, így a hívó függvénynek kell ezt megoldania.
        /// </remarks>
        public bool SendExceptionMail(DateTime date, string appName, string sender, MailAddressCollection addresses, Exception ex)
        {
            string subjectTmp, messageTmp;
            List<Attachment> attachmentsTmp = new List<Attachment>();

            subjectTmp = "Exception e-mail by " + sender + ", " + date.ToShortDateString(); // Az üzenet tárgya

            messageTmp = // Az üzenet szövege
                "Dear Sir/Madam," + Environment.NewLine +
                Environment.NewLine +
                "this e-mail is automatically sended to you by " + sender + ", please don't reply to it!" + Environment.NewLine +
                "The message was generated by the " + appName + " software." + Environment.NewLine +
                Environment.NewLine +
                "System messages:" + Environment.NewLine +
                "An exception has been occurred, which needs user interaction! The software has suspended its operation!" +
                Environment.NewLine + "The exception is the following:" + Environment.NewLine +
                ex.Message + Environment.NewLine +
                Environment.NewLine +
                "If you don't want to receive these e-mails, please contact with the system administrator." + Environment.NewLine;

            return SendMail(subjectTmp, messageTmp, attachmentsTmp, addresses); // Elküldjük a levelet
        }

        /// <summary>Ez a függvény e-mail küld a megadott hibáról</summary>
        /// <param name="date">A küldés dátuma</param>
        /// <param name="appName">A program neve (verziószámmal)</param>
        /// <param name="sender">A küldő</param>
        /// <param name="addresses">A címzettek</param>
        /// <returns>A visszatérési érték igaz, ha az e-mail elküldése sikeres volt, hamis, ha nem sikerült a küldés</returns>
        /// <remarks>
        /// Az elküldött fájl a küldés dátuma előtti nap logfájlja lesz
        /// A függvény nem rendelkezik kivételkezeléssel, így a hívó függvénynek kell ezt megoldania.
        /// </remarks>
        public bool SendFaultMail(DateTime date, string appName, string sender, MailAddressCollection addresses, string description)
        {
            string subjectTmp, messageTmp;
            List<Attachment> attachmentsTmp = new List<Attachment>();

            subjectTmp = "Fault e-mail by " + sender + ", " + date.ToShortDateString(); // Az üzenet tárgya

            messageTmp = // Az üzenet szövege
                "Dear Sir/Madam," + Environment.NewLine +
                Environment.NewLine +
                "this e-mail is automatically sended to you by " + sender + ", please don't reply to it!" + Environment.NewLine +
                "The message was generated by the " + appName + " software." + Environment.NewLine +
                Environment.NewLine +
                "System messages:" + Environment.NewLine +
                description +
                Environment.NewLine +
                "If you don't want to receive these e-mails, please contact with the system administrator." + Environment.NewLine;

            return SendMail(subjectTmp, messageTmp, attachmentsTmp, addresses); // Elküldjük a levelet
        }

        /// <summary>Ez a függvény e-mailt küld a megadott paraméterekkel</summary>
        /// <param name="subject">Az e-mail tárgya</param>
        /// <param name="message">Az e-mail szövege</param>
        /// <param name="attachments">Az e-mail mellékletei</param>
        /// <param name="addresses">Az e-mail címzettjei</param>
        /// <returns>A visszatérési érték igaz, ha az e-mail elküldése sikeres volt, hamis, ha nem sikerült a küldés</returns>
        /// <remarks>
        /// A függvény csak az e-mail küldéssel kapcsolatos kivételeket kezei, így a hívó függvénynek kell megoldania a
        /// többi kivétel kezelését.
        /// </remarks>
        private bool SendMail(string subject, string message, List<Attachment> attachments, MailAddressCollection addresses)
        {
            SmtpClient smtp = new SmtpClient("smtp.gmail.com", 587); // Az SMTP kliens az e-mailek küldéséhez
            smtp.EnableSsl = true; // SSL használata
            smtp.DeliveryMethod = SmtpDeliveryMethod.Network; // A küldési mód
            smtp.UseDefaultCredentials = false; // Nem az alapértelmezett adatokat használjuk,
            smtp.Credentials = new NetworkCredential("effekta.hungary@gmail.com", "Q19Slt25"); // hanem ezt a címet és jelszót

            MailMessage mailTmp = new MailMessage(); // A küldendő üzenet
            mailTmp.From = new MailAddress("effekta.hungary@gmail.com"); // A küldő

            IEnumerator<MailAddress> enumerator1 = addresses.GetEnumerator(); // Kiolvassuk a címzettek listáját
            while (enumerator1.MoveNext())
                mailTmp.To.Add(enumerator1.Current); // A címeket betesszük az e-mail címzettjei közé

            mailTmp.Subject = subject; // Beállítjuk az üzenet tárgyát és szövegét
            mailTmp.Body = message;

            IEnumerator<Attachment> enumerator2 = attachments.GetEnumerator(); // Kiolvassuk a csatolmányok listáját
            while (enumerator2.MoveNext())
                mailTmp.Attachments.Add(enumerator2.Current); // A csatolmányokat betesszük az e-mail mellékletei közé

            bool success = true; // Ha sikeres lesz a művelet, akkor ezzel térünk vissza

            try
            {
                if (mailTmp.To.Count > 0) // Csak akkor küldjük el a levelet, ha legalább egy címzett van a listában
                    smtp.Send(mailTmp);
                else
                {
                    success = false; // Sikertelen volt az e-mail küldés, így a hibátbeírjuk a kivétel-fájlba
                    WriteExceptionFile(DateTime.Now, "SendMail: The e-mail wasn't sended, because there is no address in the list!");
                }

                smtp.Dispose(); // Kilépés az SMTP szerverről
            }
            catch (Exception e) // Kivételkezelés
            {
                success = false; // Sikertelen volt az e-mail küldés
                WriteExceptionFile(DateTime.Now, "SendMail: " + e.Message); // A hiba kiírása a kivétel-fájlba

                smtp.Dispose(); // Kilépés az SMTP szerverről
            }

            return success;
        }
    }
}
