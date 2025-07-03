// License information
/*
TicketleoConverter can import guest lists downloaded from the
popular reservation platform and convert them into a usable
and printable spreadsheet document

    Copyright (c) 2025 BimstiBamsti

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <https://www.gnu.org/licenses/>.
*/

/**************************************
 * TicketleoConverter, main.cpp
 *
 * by BimstiBamsti - 2025
 **************************************/

#include <QAbstractButton>
#include <QApplication>
#include <QMessageBox>
#include <QFileDialog>
#include <QIcon>
#include <QPushButton>
#include <QStandardPaths>
#include <QLibraryInfo>
#include <QTranslator>

#include "xlsxdocument.h"

using namespace QXlsx;

struct Reservation
{
    int number;
    QString firstName;
    QString lastName;
    int preis;
    int count;
    QString seats;
};

int main(int argc, char* argv[])
{
    QApplication a(argc, argv);
    a.setWindowIcon(QIcon(":/icons/AppIcon"));

    QCoreApplication::setApplicationName(QString("TicketleoConverter"));

    const QString windowTitle = QString("%1 - %2").arg(
            QCoreApplication::applicationName(), COMMIT_VERSION);

    QTranslator qtBaseTranslator;
    if (!qtBaseTranslator.load(QLocale::system(), "qtbase", "_",
                               QLibraryInfo::path(QLibraryInfo::TranslationsPath)))
        qWarning("Basis-Sprachpaket konnte nicht geladen werden - Systemfehler!");
    a.installTranslator(&qtBaseTranslator);

    QMessageBox msgBox;
    msgBox.setWindowTitle(windowTitle);
    msgBox.setText(
            QString("<p><center><u><b style='font-size: large;'>%1 - %2</b></u><br>"
                    "<i style='font-size: small;'>Konvertiert eine "
                    "Ticketleo-Exportdatei in "
                    " ein verwendbares Format</i></center></p>"
                    "<p><b>Exportdatei erstellen:</b>"
                    "<ul>"
                    "<li>Bei Ticketleo einloggen</li>"
                    "<li>\"Bestellungen\" - \"Daten exportieren\" klicken</li>"
                    "<li>Vorstellung auswählen <i style='font-size: small;'>(keine "
                    "weiteren "
                    "Einstllungen nötig)</i></li>"
                    "<li>\"Export starten\" klicken</li></ul></p>"
                    "<p><b>Anwendung:</b>"
                    "<ul>"
                    "<li>\"Datei laden...\" klicken</li>"
                    "<li>Heruntergeladene Datei auswählen</li>"
                    "<li>Zieldatei angeben</li>"
                    "<li>\"Speichern\" klicken</li></ul></p>"
                    "<p><b>Fertigstellung:</b>"
                    "<ul>"
                    "<li>Erzeugte Datei öffnen</li>"
                    "<li>Druckränder kontrollieren/anpassen</li>"
                    "<li>Arbeitsblatt ausdrucken</li></ul></p>")
                    .arg(QApplication::applicationName(), COMMIT_VERSION));

    msgBox.setIconPixmap(QPixmap(":/icons/AppIcon"));
    msgBox.addButton(QMessageBox::Close);
    QPushButton* openButton = msgBox.addButton(QMessageBox::Open);
    openButton->setText("Datei laden...");

    QPushButton* helpButton = msgBox.addButton(QMessageBox::Help);
    helpButton->setText("Info");
    helpButton->disconnect();
    msgBox.connect(helpButton, &QPushButton::clicked, &msgBox, [&]() {
        QMessageBox::information(&msgBox, windowTitle,
                                 "<a href=\"https://github.com/BimstiBamsti/TicketleoConverter\">"
                                 "TicketleoConverter</a><br><br>"
                                 "<a href=\"https://github.com/QtExcel/QXlsx\">"
                                 "QXlsx</a><br><br>"
                                 "<a href=\"https://www.qt.io\">"
                                 "Qt</a><br><br>"
                                 "<a href=\"https://www.ticketleo.com\">"
                                 "Ticketleo - Event Ticketing mit Platzreservierung</a>");
    });

    if (msgBox.exec() != QMessageBox::Open)
        return 0;

#ifndef QT_DEBUG
    QStringList dlPaths = QStandardPaths::standardLocations(QStandardPaths::DownloadLocation);
#endif
    // get input file
    QString fileName = QFileDialog::getOpenFileName(
            nullptr, "Datei öffnen",
#ifdef QT_DEBUG
            QDir::homePath() + "/workspace/TicketleoConverter/",
#else
            dlPaths.isEmpty() ? QDir::homePath() : dlPaths.constFirst(),
#endif
            "Excel Dateien (*.xlsx)");

    if (fileName.isEmpty())
        return 0;

    // load and check input file
    Document inputDoc(fileName);
    QString docTitle = inputDoc.read(1, 1).toString(); // representing "A1"

    if (docTitle.isEmpty()) {
        QMessageBox::critical(nullptr, "Fehler!",
                              "Titel der Datei konnte nicht gelesen werden!\n"
                              "Dateiformat fehlerhaft.");
        return -1;
    }

    if (inputDoc.read("A3").toString() != "Reservierungsnr.") {
        QMessageBox::critical(nullptr, "Fehler!",
                              "Startmarke nicht gefunden!\n"
                              "Dateiformat fehlerhaft.");
        return -2;
    }

    // read input data
    QList<Reservation> bookingList;
    QString str;
    int index = 4;

    do {
        str = inputDoc.read(index, 1).toString();
        if (!str.isEmpty()) {
            Reservation res;
            res.number = str.toInt();
            res.firstName = inputDoc.read(index, 3).toString();
            res.lastName = inputDoc.read(index, 4).toString();
            res.preis = inputDoc.read(index, 7).toInt();
            res.count = inputDoc.read(index, 8).toInt();
            res.seats = inputDoc.read(index, 9)
                                .toString()
                                .replace("(Plätze, Sitzplatz),", "")
                                .replace("(Plätze, Sitzplatz)", "");
            bookingList.append(res);
        }
        index++;
    } while (!str.isEmpty());

    // sort bookings
    std::sort(bookingList.begin(), bookingList.end(),
              [](Reservation left, Reservation right) {
                  if (left.lastName == right.lastName) {
                      if (left.firstName == right.firstName)
                          return left.number < right.number;
                      else
                          return left.firstName < right.firstName;
                  } else
                      return left.lastName < right.lastName;
              });

    // create output document and define formats
    Document outputDoc;

    Format formatHeader;
    formatHeader.setFontName("Arial");
    formatHeader.setFontSize(10);
    formatHeader.setHorizontalAlignment(Format::AlignHCenter);
    formatHeader.setVerticalAlignment(Format::AlignVCenter);
    formatHeader.setFontBold(true);
    formatHeader.setBorderStyle(Format::BorderThin);

    Format formatLeftTop;
    formatLeftTop.setFontName("Arial");
    formatLeftTop.setFontSize(10);
    formatLeftTop.setHorizontalAlignment(Format::AlignLeft);
    formatLeftTop.setVerticalAlignment(Format::AlignTop);
    formatLeftTop.setBorderStyle(Format::BorderHair);

    Format formatNr;
    formatNr.setFontName("Arial");
    formatNr.setFontSize(10);
    formatNr.setHorizontalAlignment(Format::AlignHCenter);
    formatNr.setVerticalAlignment(Format::AlignTop);
    formatNr.setBorderStyle(Format::BorderHair);

    Format formatSeats;
    formatSeats.setFontName("Arial");
    formatSeats.setFontSize(10);
    formatSeats.setVerticalAlignment(Format::AlignTop);
    formatSeats.setTextWrap(true);
    formatSeats.setBorderStyle(Format::BorderHair);

    Format formatSummary;
    formatSummary.setFontName("Arial");
    formatSummary.setFontSize(10);
    formatSummary.setVerticalAlignment(Format::AlignVCenter);
    formatSummary.setFontItalic(true);

    Format formatAbout;
    formatAbout.setFontName("Arial");
    formatAbout.setFontSize(8);
    formatAbout.setHorizontalAlignment(Format::AlignRight);
    formatAbout.setVerticalAlignment(Format::AlignVCenter);
    formatAbout.setFontItalic(true);

    // write column header
    outputDoc.write(1, 1, "Nummer", formatHeader);
    outputDoc.write(1, 2, "Vorname", formatHeader);
    outputDoc.write(1, 3, "Nachname", formatHeader);
    outputDoc.write(1, 4, "Anzahl", formatHeader);
    outputDoc.write(1, 5, "Sitzplätze", formatHeader);
    outputDoc.write(1, 6, "Preis", formatHeader);

    // write data and set row hight
    int total = 0;
    int line = 2;
    for (const auto &res : std::as_const(bookingList)) {
        outputDoc.write(line, 1, res.number, formatLeftTop);
        outputDoc.write(line, 2, res.firstName, formatLeftTop);
        outputDoc.write(line, 3, res.lastName, formatLeftTop);
        outputDoc.write(line, 4, res.count, formatNr);
        outputDoc.write(line, 5, res.seats, formatSeats);
        outputDoc.write(line, 6, res.preis, formatNr);

        double h = (res.count * (formatLeftTop.fontSize() * 1.13)) + 2.0;
        outputDoc.setRowHeight(line, h);

        total += res.count;
        line++;
    }

    // write summary and version
    outputDoc.setRowHeight(line++, 5); // set space
    outputDoc.write(line, 1,
                    QString("Buchungen: %1, Reservierte Plätze: %2 --- %3")
                            .arg(bookingList.length())
                            .arg(total)
                            .arg(QDateTime::currentDateTime().toString("d.M.yyyy hh:mm")),
                    formatSummary);

    outputDoc.write(line, 6, windowTitle, formatAbout);

    // set pagemargin and column widths
    outputDoc.currentWorksheet()->setPageMargin(0.5, 0.5, 0.3, 0.8, 0.2, 0.15); // inch

    outputDoc.setColumnWidth(1, 8);
    outputDoc.setColumnWidth(2, 15);
    outputDoc.setColumnWidth(3, 20);
    outputDoc.setColumnWidth(4, 6);
    outputDoc.setColumnWidth(5, 15);
    outputDoc.setColumnWidth(6, 5);

    // set header and footer
    outputDoc.currentWorksheet()->writeHeader(
            QString("&L&\"Arial,Standard\"&10Reservierungen %1").arg(docTitle));
    outputDoc.currentWorksheet()->writeFooter("&C&\"Arial,Standard\"&10&P/&N");

    // save file
    QString saveFileName = QFileDialog::getSaveFileName(
            nullptr, "Datei speichern",
#ifdef QT_DEBUG
            QDir::homePath() + "/workspace/TicketleoConverter/" + QString("%1.xlsx").arg(docTitle),
#else
            QDir::homePath() + QString("/%1.xlsx").arg(docTitle),
#endif
            "Excel Dateien (*.xlsx)");

    if (saveFileName.isEmpty())
        return 0;

    if (!outputDoc.saveAs(saveFileName)) {
        QMessageBox::critical(nullptr, "Fehler!",
                              "Datei konnte nicht gespeichert werden!");
        return -3;
    }

    QMessageBox::information(nullptr, windowTitle,
                             QString("Datei wurde gespeichert:\n"
                                     "'%1'\n\n"
                                     "Titel: '%2'\n"
                                     "Buchungen: %3\n"
                                     "Plätze: %4")
                                     .arg(saveFileName, docTitle)
                                     .arg(bookingList.length())
                                     .arg(total));

    return 0;
}
