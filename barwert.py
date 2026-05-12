# coding: utf-8
"""
Copyright (C) 2026 by Jörg Höttges, joerg.hoettges@posteo.de

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <https://www.gnu.org/licenses/>.
"""

import os
import csv
import sys
import tkinter

commandlis = ['Jahr', '1. Rate', 'Laufzeit', 'Vorlaufzeit', 'Abnahme pro Jahr', 
'Zunahme pro Jahr', 'Zahltakt', 'Quadratische Steigerung', 'Steigerungsrate', 'Einmalzahlung',]
commandbas = ['Gesamtlaufzeit', 'Abzinsung', ]


class KVR():
    """Verwaltet Teilzahlungen und Barwerte der Teilkostenreihen"""
    def __init__(self, fw):
        """Aufgabebuffer und internen Clipboardbuffer initialisiseren"""
        self.clipboard = ''     # enthält zuletzt übernommenes Clipboard
        self.separator = '\t'

    def reihe(self, comlis):
        """Berechnung des Barwertes für eine Teilkostenreihe"""
        jahr = comlis.get('Jahr', 0)
        rate0 = float(comlis.get('Einmalzahlung', '0').replace(',', '.'))
        rate0 = float(comlis.get('1. Rate', '0').replace(',', '.'))
        ngesamt = int(comlis.get('Gesamtlaufzeit', '1').replace(',', '.'))
        nlauf = int(comlis.get('Laufzeit', str(ngesamt)).replace(',', '.'))
        nvor = int(comlis.get('Vorlaufzeit', '0').replace(',', '.'))
        d1 = float(comlis.get('Abnahme pro Jahr', '0').replace(',', '.'))
        d2 = float(comlis.get('Zunahme pro Jahr', '0').replace(',', '.'))
        takt = int(comlis.get('Zahltakt', '1').replace(',', '.'))
        deg1 = d2 - d1
        deg2 = float(comlis.get('Quadratische Steigerung', '0').replace(',', '.'))

        t = comlis.get('Abzinsung', '0').replace(',', '.')
        if t.count(';') > 1:
            # Es wurde eine Liste mit Prozentwerten übergeben
            abzinsung = [float(el.replace('%', '').replace(',', '.'))/100. for el in t.split(';')]
            nfill = ngesamt - len(abzinsung)                            # fehlende Werte bis ngesamt
            abzinsung.extend([0.] * nfill)
        elif "%" in t:
            abzinsung = [float(t.replace('%', ''))/100.] * ngesamt

        t = comlis.get('Steigerungsrate', '0').replace(',', '.')
        steigrate = float(t.replace('%', ''))/100.

        kap = 0.
        fak = 1.
        abzins = 1.
        for i in range(1, 1 + ngesamt):
            abzins = (1. + abzinsung[i - 1]) ** i
            if nvor + nlauf >= i > nvor:
                fak *= 1 + steigrate
                if (i - nvor) % takt == 0:
                    teilbetrag = (rate0 + deg1 * (i - nvor) + deg2 * (i - nvor)**2) * fak / abzins
                    # fw.write(f'{i=}: {teilbetrag=}, {fak=}, {abzins=}\n')
                    kap += (rate0 + deg1 * (i - nvor) + deg2 * (i - nvor)**2) * fak / abzins
        return kap

    def update(self):
        """Liest Clipboard"""
        clipboard = form.clipboard_get()

        # Initialisierung
        kap: float = 0.                 # Barwert für alle Teilkostenreihen
        teilkap:List[float] = []        # Liste der Barwerte der Teilkostenreihen
        titles:List[str] = []
        comlis:dict[str, str] = {}      # Parameter der Teilkostenreihe 
        header = None                   # Merker, dass nächste zu verarbeitende Zeile Spaltenbezeichnungen enthält
        commonpart = True               # Merker, dass noch Allgemeine Parameter gelesen werden müssen.

        lines = clipboard.split('\n')
        for line in lines:
            if header is None:
                # Einlesen der Spaltenbezeichnungen für Allgemeine Parameter sowie Kostenreihen
                header = line.split('\t')
                # Leerzeilen überspringen
                if not any([titel in commandbas + commandlis for titel in header]):
                    header = None
            elif commonpart:
                # Einlesen der Allgemeinen Parameter
                vals = line.split('\t')
                comlis = {key: val for key, val in zip(header, vals) if key in commandbas}
                # Merker setzen zur Verarbeitung der Teilkostenreihen
                commonpart = False
                header = None
            else:
                # Einlesen der Kostenreihen. Anschließend Teilberechnungen
                if line.replace('\t', '').strip() == '':
                    continue
                vals = line.split('\t')
                comlis |= {key: val for key, val in zip(header, vals) if key in commandbas + commandlis and val.strip() != ''}
                title = '\t'.join([val for key, val in zip(header, vals) if key not in commandbas + commandlis])

                # Berechnung des Barwertes für Teilkostenreihe
                erg = self.reihe(comlis)
                args = '\n'.join([f'{key}: {comlis[key]}' for key in comlis])
                fw.write(f'{title}\n{args}\nTeilbarwert: {erg}\n\n')
                titles.append(title)
                teilkap.append(erg)
                
                # Vorbereitung für nächste Kostenreihe: Parameter auf allgemeine Parameter zurücksetzen
                comlis = {'Abzinsung': comlis['Abzinsung'], 'Gesamtlaufzeit': comlis['Gesamtlaufzeit']}
        if 'erg' not in vars():
            return 0, [0], ['Fehler: Kopierte Daten konnten nicht verarbeitet werden!']
        return erg, teilkap, titles
                
def calc_barwert():
    erg, teilkap, titles = kvr.update() 
    kap = sum(teilkap)
    form.clipboard_clear()
    clipboard = f'Barwert:\t{kap:.2f}'.replace('.', ',')
    form.clipboard_append(clipboard)
    textline.config(text = 'Berechnung des Barwertes ...')

    form.destroy()

def calc_einzelbarwerte():
    erg, teilkap, titles = kvr.update()
    kap = sum(teilkap)
    form.clipboard_clear()
    clipboard = '\n'.join([f'{title}:\t{kap:.2f}'.replace('.', ',') for title, kap in zip(titles, teilkap)])

    n = len(titles[0].split('\t'))
    title = '\t'*(n-1) + 'Barwert'
    clipboard += f'\n\n{title}:\t{kap:.2f}'.replace('.', ',')
    form.clipboard_append(clipboard)
    textline.config(text = 'Berechnung der Einzelbarwerte ...')

    form.destroy()

dir = os.environ['HOMEPATH']
drive = os.environ['HOMEDRIVE']
filenam = os.path.join(drive, dir, 'Documents/kvr.txt')
with open(filenam, 'w') as fw:

    kvr = KVR(fw)

    form = tkinter.Tk()
    form.title('KVR-Leitlinien: Berechnung des Barwertes')
    form.geometry('600x200')

    textline = tkinter.Label(
        form,
        text='     Kopieren Sie zuerst die Teilzahlungsvorgänge aus der Excel-Datei ...\n',
        anchor="w",
        justify="left",
        fg =    'Blue',
        font =  'lucida 12',
        height = 2,
        width = 70,
    )
    textline.pack(pady=10)

    # bu_barwert = tkinter.Button(
        # form, 
        # text=   'Barwert berechnen und in Zwischenablage übernehmen', 
        # fg =    'Black',
        # font =  'lucida 12',
        # width = 50,
        # command=calc_barwert,
    # )
    # bu_barwert.pack(pady=10)

    bu_barwertliste = tkinter.Button(
        form,
        text=   'Barwerte berechnen und in Zwischenablage übernehmen', 
        fg =    'Black',
        font =  'lucida 12',
        height = 1,
        width = 50,
        command=calc_einzelbarwerte,
    )
    bu_barwertliste.pack(pady=10)

    textline = tkinter.Label(
        form,
        text='     Ergebnis in Datei: Dokumente\\kvr.txt',
        anchor="w",
        justify="left",
        fg =    'Blue',
        font =  'lucida 12',
        height = 2,
        width = 70,
    )
    textline.pack(pady=10)

    form.bind('<Escape>', lambda e, w=form: form.destroy())

    form.mainloop()

del fw
