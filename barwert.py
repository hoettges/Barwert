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

import matplotlib.pyplot as plt
import csv
import sys
import tkinter

commandlis = ['Jahr', '1. Rate', 'Laufzeit', 'Vorlaufzeit', 'Abnahme pro Jahr', 
'Zunahme pro Jahr', 'Zahltakt', 'Quadratische Steigerung', 'Steigerungsrate', 'Einmalzahlung',]
commandbas = ['Gesamtlaufzeit', 'Inflation', ]

class KVR():
    def __init__(self):
        self.barwerte = []       # Enthält Barwerte der Teilzahlungen
        self.commands = None   # enthält nach Übernahme aus Clipboard Teilzahlungsliste als Titel, Commands und Werte
        self.separator = '\t'

    def reihe(self, comlis):
        jahr = comlis.get('Jahr', 0)
        rate0 = float(comlis.get('Einmalzahlung', 0))
        rate0 = float(comlis.get('1. Rate', 0))
        ngesamt = int(comlis.get('Gesamtlaufzeit', 1))
        nlauf = int(comlis.get('Laufzeit', ngesamt))
        nvor = int(comlis.get('Vorlaufzeit', 0))
        d1 = float(comlis.get('Abnahme pro Jahr', 0))
        d2 = float(comlis.get('Zunahme pro Jahr', 0))
        takt = int(comlis.get('Zahltakt', 1))
        deg1 = d2 - d1
        deg2 = float(comlis.get('Quadratische Steigerung', 0))

        t = comlis.get('Inflation', 0)
        if isinstance(t, str):
            if "%" in t:
                inflation = [float(t.replace('%', ''))/100.] * ngesamt
            elif t.count(';') > 2:
                # Es wurde eine Liste mit Prozentwerten (!) übergeben
                inflation = [float(el.replace(',', '.')) for el in t.split(';')]
                nfill = ngesamt - len(inflation)                            # fehlende Werte bis ngesamt
                inflation.extend([0] * nfill)
        else:
            inflation = [float(t)/100.] * ngesamt

        t = comlis.get('Steigerungsrate', 0)
        if isinstance(t, str):
            steigrate = float(t.replace('%', ''))/100.
        else:
            steigrate = float(t)/100.

        kap = 0
        fak = 1
        abzins = 1
        for i in range(1, 1 + ngesamt):
            abzins = (1 + inflation[i - 1]) ** i
            if nvor + nlauf >= i > nvor:
                fak *= 1 + steigrate
                if (i - nvor) % takt == 0:
                    teilbetrag = (rate0 + deg1 * (i - nvor) + deg2 * (i - nvor)**2) * fak / abzins
                    print(f'{i=}: {teilbetrag=}, {fak=}, {abzins=}')
                    kap += (rate0 + deg1 * (i - nvor) + deg2 * (i - nvor)**2) * fak / abzins
        return kap

    def calc(self, commands):
        kap = 0
        comlis = {}
        kopf = True
        zeile = ''
        teilkap = []
        for zeile in commands:
            if zeile.count(self.separator) == 0 or zeile.strip() == self.separator or zeile.startswith('Allgemeine'):
                continue
            # if zeile.count(self.separator) < 2:
                # zeile += zeil_
                # if zeile.count(self.separator) < 2:
                    # continue
            # else:
                # zeile = zeil_
            print(f'{zeile=}')
            command, value = zeile.split(self.separator)
            if command not in commandlis:
                if not kopf:
                    # Berechnung des Zahlungsprozesses
                    if comlis.get('Jahr') is None and len(comlis) > 1:
                        raise BaseException('Fehler: Jahr muss immer gegeben sein!')
                    elif comlis.get('1. Rate') is None and len(comlis) > 1:
                        raise BaseException('Fehler: Betrag oder 1. Rate muss immer gegeben sein!')
                    elif comlis.get('Inflation') is None and len(comlis) > 1:
                        raise BaseException('Fehler: Inflation muss immer gegeben sein!')
                    elif comlis.get('Gesamtlaufzeit') is None and len(comlis) > 1:
                        raise BaseException('Fehler: Gesamtlaufzeit muss immer gegeben sein!')
                    erg = self.reihe(comlis)
                    teilkap.append(erg)
                    # Parameter zurücksetzen
                    comlis = {'Inflation': comlis['Inflation'], 'Gesamtlaufzeit': comlis['Gesamtlaufzeit']}
                    kopf = True
            else:
                kopf = False
            if command in commandbas + commandlis:
                comlis[command] = value.replace(',', '.')
                print(f'{command=}, {value=}')
        # print(f'{comlis=}')
        erg = self.reihe(comlis)
        teilkap.append(erg)
        return teilkap

class Buffer():
    """Verwaltet Teilzahlungen und Barwertliste"""
    def __init__(self):
        """Aufgabebuffer und internen Clipboardbuffer initialisiseren"""
        self.clipboard = ''     # enthält zuletzt übernommenes Clipboard
        self.separator = '\t'
        self.commands = []

    def update(self):
        """Prüft Stand des internen Clipboards und übernimmt bei Änderung den neuen Stand"""
        clipnew = form.clipboard_get()
        if any(clipnew.startswith(t) for t in ['Allgemeine', 'Inflation', 'Gesamtlaufzeit']):
            # Testen auf Seperator 
            # if clipnew.count(';') > clipnew.count('\t'):
                # self.separator = ';'
            self.clipboard = clipnew
            self.commands = clipnew.split('\n')
        return self.commands

def calc_barwert():
    commands = buf.update()
    kvr.separator = buf.separator
    teilkap = kvr.calc(commands)
    kap = sum(teilkap)
    clipboard = f'Barwert:\t{kap:.2f}'.replace('.', ',')
    form.clipboard_clear()
    form.clipboard_append(clipboard)
    textline.config(text = 'Übernehmen Sie nun den Barwert aus der Zwischenablage ...')

def calc_einzelbarwerte():
    commands = buf.update()
    kvr.separator = buf.separator
    teilkap = kvr.calc(commands)
    clipboard = '\n'.join([f'Barwert:\t{kap:.2f}'.replace('.', ',') for kap in teilkap])
    form.clipboard_clear()
    form.clipboard_append(clipboard)
    textline.config(text = 'Übernehmen Sie nun die Einzelbarwerte aus der Zwischenablage ...')

kvr = KVR()
buf = Buffer()

form = tkinter.Tk()
form.title('KVR-Leitlinien: Berechnung des Barwertes')
form.geometry('600x175')
textline = tkinter.Label(
    form, text='Kopieren Sie zuerst die Teilzahlungsvorgänge aus der Excel-Datei ...',
    fg =    'Blue',
    font =  'lucida 12',
)
textline.pack(pady=10)

bu_barwert = tkinter.Button(
    form, 
    text=   'Berechneten Barwert in Zwischenablage übernehmen', 
    fg =    'Black',
    font =  'lucida 12',
    width = 50,
    command=calc_barwert,
)
bu_barwert.pack(pady=10)
bu_barwertliste = tkinter.Button(
    form,
    text=   'Berechnete Einzelbarwerte in Zwischenablage übernehmen', 
    fg =    'Black',
    font =  'lucida 12',
    width = 50,
    command=calc_einzelbarwerte,
)
bu_barwertliste.pack(pady=10)

form.mainloop()
