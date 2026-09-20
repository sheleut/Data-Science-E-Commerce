"""Erzeugt die Praesentation aus den Ergebnissen des Notebooks."""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

BLAU = RGBColor(0x2A, 0x78, 0xD6)
ORANGE = RGBColor(0xEB, 0x68, 0x34)
INK = RGBColor(0x0B, 0x0B, 0x0B)
GRAU = RGBColor(0x52, 0x51, 0x4E)
HELLGRAU = RGBColor(0x89, 0x87, 0x81)
WEISS = RGBColor(0xFF, 0xFF, 0xFF)

prs = Presentation()
prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
BREITE, HOEHE = prs.slide_width, prs.slide_height
RAND = Inches(0.9)
INHALT_BREITE = BREITE - 2 * RAND


def leere_folie():
    return prs.slides.add_slide(prs.slide_layouts[6])


def textfeld(folie, links, oben, breite, hoehe):
    tf = folie.shapes.add_textbox(links, oben, breite, hoehe).text_frame
    tf.word_wrap = True
    return tf


def absatz(tf, text, groesse=18, fett=False, farbe=INK, erster=False,
           abstand_vor=6, zeilenabstand=1.25):
    p = tf.paragraphs[0] if erster else tf.add_paragraph()
    p.text = text
    p.space_before = Pt(0 if erster else abstand_vor)
    p.line_spacing = zeilenabstand
    for r in p.runs:
        r.font.size, r.font.bold, r.font.color.rgb = Pt(groesse), fett, farbe
        r.font.name = "Calibri"
    return p


def kopfzeile(folie, titel, unterzeile=None):
    """Titelbereich + duenne Trennlinie."""
    tf = textfeld(folie, RAND, Inches(0.45), INHALT_BREITE, Inches(0.9))
    absatz(tf, titel, groesse=30, fett=True, erster=True, zeilenabstand=1.0)
    linie = folie.shapes.add_shape(1, RAND, Inches(1.32), INHALT_BREITE, Pt(2.5))
    linie.fill.solid(); linie.fill.fore_color.rgb = BLAU
    linie.line.fill.background(); linie.shadow.inherit = False
    if unterzeile:
        tf2 = textfeld(folie, RAND, Inches(1.45), INHALT_BREITE, Inches(0.5))
        absatz(tf2, unterzeile, groesse=15, farbe=GRAU, erster=True, zeilenabstand=1.1)
    return Inches(2.05) if unterzeile else Inches(1.75)


def kachel(folie, links, oben, breite, hoehe, zahl, beschriftung, farbe=BLAU):
    box = folie.shapes.add_shape(1, links, oben, breite, hoehe)
    box.fill.solid(); box.fill.fore_color.rgb = RGBColor(0xF4, 0xF7, 0xFB)
    box.line.color.rgb = RGBColor(0xE1, 0xE0, 0xD9); box.line.width = Pt(1)
    box.shadow.inherit = False
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.15)
    tf.margin_top = Inches(0.18)
    absatz(tf, zahl, groesse=34, fett=True, farbe=farbe, erster=True, zeilenabstand=1.0)
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    p = absatz(tf, beschriftung, groesse=13, farbe=GRAU, abstand_vor=4, zeilenabstand=1.15)
    p.alignment = PP_ALIGN.CENTER


def bild_einpassen(folie, pfad, links, oben, max_breite, max_hoehe):
    from PIL import Image
    with Image.open(pfad) as im:
        bb, bh = im.size
    faktor = min(max_breite / bb, max_hoehe / bh)
    b, h = int(bb * faktor), int(bh * faktor)
    folie.shapes.add_picture(pfad, links + int((max_breite - b) / 2),
                             oben + int((max_hoehe - h) / 2), width=b, height=h)


def tabelle(folie, daten, links, oben, breite, hoehe, spaltenbreiten=None,
            hervorheben=None):
    zeilen, spalten = len(daten), len(daten[0])
    shape = folie.shapes.add_table(zeilen, spalten, links, oben, breite, hoehe)
    tab = shape.table
    if spaltenbreiten:
        for i, b in enumerate(spaltenbreiten):
            tab.columns[i].width = b
    for r in range(zeilen):
        for c in range(spalten):
            zelle = tab.cell(r, c)
            zelle.text = str(daten[r][c])
            zelle.margin_left = zelle.margin_right = Inches(0.1)
            zelle.margin_top = zelle.margin_bottom = Inches(0.04)
            p = zelle.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT if c == 0 else PP_ALIGN.CENTER
            for run in p.runs:
                run.font.size = Pt(15 if r == 0 else 14)
                run.font.name = "Calibri"
                run.font.bold = (r == 0)
                run.font.color.rgb = WEISS if r == 0 else INK
            zelle.fill.solid()
            if r == 0:
                zelle.fill.fore_color.rgb = BLAU
            elif hervorheben and r in hervorheben:
                zelle.fill.fore_color.rgb = RGBColor(0xFD, 0xF0, 0xE9)
            else:
                zelle.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF) if r % 2 else RGBColor(0xF7, 0xF7, 0xF5)
    return shape


def notizen(folie, zeit, text):
    """Sprechernotiz mit Zeitbudget."""
    folie.notes_slide.notes_text_frame.text = f"[{zeit}]\n\n{text}"


def fusszeile(folie, nummer):
    tf = textfeld(folie, RAND, HOEHE - Inches(0.62), INHALT_BREITE, Inches(0.35))
    p = absatz(tf, f"E-Commerce Lieferprognose  |  {nummer}", groesse=11,
               farbe=HELLGRAU, erster=True)
    p.alignment = PP_ALIGN.RIGHT


B = "praesentation/bilder/"

# ---------------------------------------------------------------- Titelfolie
f = leere_folie()
balken = f.shapes.add_shape(1, 0, 0, Inches(0.28), HOEHE)
balken.fill.solid(); balken.fill.fore_color.rgb = BLAU
balken.line.fill.background(); balken.shadow.inherit = False

tf = textfeld(f, Inches(1.1), Inches(2.5), Inches(11), Inches(2.4))
absatz(tf, "Kommt die Lieferung pünktlich an?", groesse=46, fett=True, erster=True,
       zeilenabstand=1.05)
absatz(tf, "Vorhersage von Lieferverspätungen im E-Commerce", groesse=22, farbe=GRAU,
       abstand_vor=16, zeilenabstand=1.2)
notizen(f, "0:00 - 0:30  |  30 Sekunden", """Kurz begruessen und die Leitfrage stellen.

"Jeder kennt das: Man bestellt etwas online und wartet. Meine Frage war, ob man schon
beim Bestelleingang vorhersagen kann, ob eine Lieferung zu spaet kommt."

Nicht laenger aufhalten - direkt zur naechsten Folie.""")

# ------------------------------------------------- 1. Datensatz & Fragestellung
f = leere_folie()
y = kopfzeile(f, "1. Datensatz und Business-Fragestellung")

kb, luecke = Inches(2.75), Inches(0.28)
ky = y + Inches(0.35)
for i, (zahl, text, farbe) in enumerate([
        ("10.999", "Bestellungen", BLAU),
        ("11", "Merkmale", BLAU),
        ("59,7 %", "kamen zu spät", ORANGE),
        ("0", "fehlende Werte", BLAU)]):
    kachel(f, RAND + i * (kb + luecke), ky, kb, Inches(1.2), zahl, text, farbe)

box = f.shapes.add_shape(1, RAND, ky + Inches(1.75), INHALT_BREITE, Inches(1.5))
box.fill.solid(); box.fill.fore_color.rgb = RGBColor(0xF4, 0xF7, 0xFB)
box.line.color.rgb = BLAU; box.line.width = Pt(1.5); box.shadow.inherit = False
tf = box.text_frame; tf.word_wrap = True
tf.margin_left = tf.margin_right = Inches(0.35); tf.margin_top = Inches(0.28)
absatz(tf, "Business-Fragestellung", groesse=15, fett=True, farbe=BLAU, erster=True)
absatz(tf, "Lässt sich frühzeitig erkennen, ob eine Sendung sich verspätet?",
       groesse=26, fett=True, abstand_vor=10, zeilenabstand=1.2)

tf = textfeld(f, RAND, ky + Inches(3.5), INHALT_BREITE, Inches(0.6))
absatz(tf, "Binäre Klassifikation   ·   0 = pünktlich, 1 = verspätet",
       groesse=17, farbe=GRAU, erster=True)
notizen(f, "0:30 - 2:00  |  90 Sekunden", """Datensatz vorstellen: internationaler
Elektronikhaendler, 10.999 Bestellungen, 11 Merkmale - Gewicht, Versandart, Rabatt,
Kundendaten.

Die wichtigste Zahl auf dieser Folie ist 59,7 Prozent: So viele Lieferungen kamen zu
spaet. Verspaetung ist hier also der Normalfall, nicht die Ausnahme.

Diese Zahl merken - sie wird spaeter zur Messlatte fuer das Modell.

Datenqualitaet kurz erwaehnen: keine fehlenden Werte, keine Duplikate - ungewoehnlich
sauber.

Dann die Business-Frage vorlesen und festhalten: binaere Klassifikation.""")
fusszeile(f, "1 / 4")

# --------------------------------------------------------------- 2. EDA (a)
f = leere_folie()
y = kopfzeile(f, "2. Erkenntnisse aus der Datenanalyse",
              "Der Rabatt ist der stärkste Hinweis auf eine Verspätung")

bild_einpassen(f, B + "boxplot.png", RAND, y, Inches(5.5), Inches(4.3))

tx = RAND + Inches(6.0)
tb = BREITE - tx - RAND
tf = textfeld(f, tx, y + Inches(0.3), tb, Inches(1.2))
absatz(tf, "Ø Rabatt", groesse=17, farbe=GRAU, erster=True)
absatz(tf, "pünktlich  5,6 %", groesse=24, fett=True, abstand_vor=8, zeilenabstand=1.15)
absatz(tf, "verspätet  18,7 %", groesse=24, fett=True, farbe=ORANGE, abstand_vor=4,
       zeilenabstand=1.15)

tabelle(f, [["Rabatt", "Anzahl", "verspätet"],
            ["0 – 10 %", "8.352", "47 %"],
            ["über 10 %", "2.647", "100 %"]],
        tx, y + Inches(2.1), tb, Inches(1.1),
        spaltenbreiten=[Emu(int(tb * 0.36)), Emu(int(tb * 0.30)), Emu(int(tb * 0.34))],
        hervorheben={2})

tf = textfeld(f, tx, y + Inches(3.45), tb, Inches(0.9))
absatz(tf, "Eine harte Schwelle, keine Steigung.", groesse=20, fett=True, farbe=ORANGE,
       erster=True, zeilenabstand=1.25)
notizen(f, "2:00 - 4:00  |  2 Minuten  (KERNFOLIE)", """Das ist die wichtigste Folie -
hier Zeit lassen.

Zuerst den Boxplot erklaeren: Kasten = mittlere 50 Prozent, Strich = Median. Links die
puenktlichen, rechts die verspaeteten Lieferungen.

Der Unterschied ist sofort sichtbar: 5,55 Prozent gegenueber 18,66 Prozent Rabatt -
mehr als das Dreifache.

Dann die Tabelle: Es ist kein gleichmaessiger Anstieg, sondern eine harte Kante. Unter
10 Prozent Rabatt sind knapp die Haelfte verspaetet. Ueber 10 Prozent: alle. Nicht die
meisten - alle 2.647.

Hier eine kurze Pause machen, die Zahl wirken lassen.

Falls Frage kommt "warum?": Wir wissen es nicht - Korrelation ist keine Kausalitaet.
Moegliche Erklaerung waeren Rabattaktionen mit hohem Bestellaufkommen.""")
fusszeile(f, "2 / 4")

# --------------------------------------------------------------- 2. EDA (b)
f = leere_folie()
y = kopfzeile(f, "2. Erkenntnisse aus der Datenanalyse",
              "Beim Gewicht: ein kritisches Fenster in der Mitte")

bild_einpassen(f, B + "korrelation.png", RAND, y, Inches(5.3), Inches(4.3))

tx = RAND + Inches(5.9)
tb = BREITE - tx - RAND
tf = textfeld(f, tx, y + Inches(0.3), tb, Inches(1.6))
absatz(tf, "1 – 2 kg", groesse=17, farbe=GRAU, erster=True)
absatz(tf, "68 % verspätet", groesse=22, fett=True, abstand_vor=2, zeilenabstand=1.1)
absatz(tf, "2 – 4 kg", groesse=17, farbe=GRAU, abstand_vor=12)
absatz(tf, "~100 % verspätet", groesse=22, fett=True, farbe=ORANGE, abstand_vor=2,
       zeilenabstand=1.1)
absatz(tf, "4 – 6 kg", groesse=17, farbe=GRAU, abstand_vor=12)
absatz(tf, "43 % verspätet", groesse=22, fett=True, abstand_vor=2, zeilenabstand=1.1)

tf = textfeld(f, tx, y + Inches(3.3), tb, Inches(1.2))
absatz(tf, "Zwei neue Merkmale", groesse=17, fett=True, farbe=BLAU, erster=True)
absatz(tf, "High_discount        0,40 → 0,46", groesse=16, abstand_vor=8)
absatz(tf, "Weight_critical     –0,27 → 0,36", groesse=16, abstand_vor=4)
notizen(f, "4:00 - 5:30  |  90 Sekunden", """Die Korrelationsmatrix nur kurz zeigen -
nicht Zelle fuer Zelle erklaeren. Wichtig ist die unterste Zeile: Rabatt 0,40, Gewicht
minus 0,27, alles andere nahe null.

Der eigentliche Punkt dieser Folie ist die Lehre daraus: Beim Gewicht gibt es ein
kritisches Fenster zwischen 2 und 4 Kilo, wo praktisch alles zu spaet kommt. Leichtere
und schwerere Pakete kommen deutlich puenktlicher an.

Die Korrelation kann so etwas prinzipiell nicht zeigen, weil sie nur nach geraden
Linien sucht. Die beiden Haelften heben sich gegenseitig auf.

Deshalb habe ich zwei Ja/Nein-Merkmale gebaut. Ergebnis: Beim Gewicht steigt der
Zusammenhang von minus 0,27 auf 0,36.

Kernaussage: Erst die Daten anschauen, dann Merkmale bauen - nicht umgekehrt.""")
fusszeile(f, "2 / 4")

# ------------------------------------------------ 3. Modellwahl & Ergebnisse
f = leere_folie()
y = kopfzeile(f, "3. Modellwahl und Ergebnisse",
              "Zwei bewusst unterschiedliche Ansätze")

tf = textfeld(f, RAND, y + Inches(0.2), Inches(5.6), Inches(2.0))
absatz(tf, "Logistic Regression", groesse=22, fett=True, farbe=BLAU, erster=True)
absatz(tf, "ein Punktesystem — einfach, nachvollziehbar", groesse=17, farbe=GRAU,
       abstand_vor=6, zeilenabstand=1.2)
absatz(tf, "Random Forest", groesse=22, fett=True, farbe=BLAU, abstand_vor=24)
absatz(tf, "300 Bäume stimmen ab — findet Schwellen selbst", groesse=17, farbe=GRAU,
       abstand_vor=6, zeilenabstand=1.2)

tx = RAND + Inches(6.2)
tb = BREITE - tx - RAND
tf = textfeld(f, tx, y, tb, Inches(0.4))
absatz(tf, "2.200 Testbestellungen", groesse=16, fett=True, erster=True)
tabelle(f, [["Metrik", "LogReg", "R. Forest", "Baseline"],
            ["Accuracy", "0,672", "0,678", "0,597"],
            ["Precision", "0,997", "0,962", "0,597"],
            ["Recall", "0,452", "0,479", "1,000"],
            ["F1-Score", "0,622", "0,640", "0,747"]],
        tx, y + Inches(0.55), tb, Inches(2.1),
        spaltenbreiten=[Emu(int(tb * 0.31)), Emu(int(tb * 0.23)),
                        Emu(int(tb * 0.23)), Emu(int(tb * 0.23))],
        hervorheben={3})

tf = textfeld(f, RAND, y + Inches(3.1), INHALT_BREITE, Inches(1.4))
absatz(tf, "Precision 0,96", groesse=22, fett=True, erster=True)
absatz(tf, "warnt es, stimmt es", groesse=17, farbe=GRAU, abstand_vor=2)

tf = textfeld(f, RAND + Inches(4.0), y + Inches(3.1), Inches(4.0), Inches(1.4))
absatz(tf, "Recall 0,48", groesse=22, fett=True, farbe=ORANGE, erster=True)
absatz(tf, "findet nur die Hälfte", groesse=17, farbe=GRAU, abstand_vor=2)

tf = textfeld(f, RAND + Inches(7.6), y + Inches(3.1), Inches(4.0), Inches(1.4))
absatz(tf, "684", groesse=22, fett=True, farbe=ORANGE, erster=True)
absatz(tf, "Verspätungen übersehen", groesse=17, farbe=GRAU, abstand_vor=2)
notizen(f, "5:30 - 7:30  |  2 Minuten  (KERNFOLIE)", """Modellwahl in einem Satz je
Modell - bewusst zwei unterschiedliche Ansaetze: ein einfaches lineares und ein
komplexes Baum-Modell.

Dann zu den Metriken. Diese unbedingt erklaeren, nicht nur vorlesen:

- Accuracy: Wie viele Vorhersagen waren insgesamt richtig?
- Precision: Wenn das Modell warnt - stimmt es dann?
- Recall: Von allen echten Verspaetungen - wie viele hat es gefunden?

Der entscheidende Gegensatz: Precision 0,96, aber Recall nur 0,48.

Bild dazu: Das Modell ist wie ein Wachmann, der nur bei Einbruechen Alarm schlaegt, die
er selbst sieht. Schlaegt er Alarm, hat er recht - aber er verschlaeft die Haelfte.

Ehrlich bleiben: Die Baseline-Spalte ganz rechts ist stures Raten "immer zu spaet". Bei
Recall und F1 schlaegt sie unser Modell. Nur bei Accuracy sind wir besser. Das zeigt:
Das Modell ist zu vorsichtig eingestellt - dazu auf der letzten Folie mehr.""")
fusszeile(f, "3 / 4")

# ---------------------------------------------------- 3b. Wichtigste Merkmale
f = leere_folie()
y = kopfzeile(f, "3. Modellwahl und Ergebnisse",
              "Woran das Modell seine Entscheidung festmacht")

bild_einpassen(f, B + "importance.png", RAND, y, Inches(6.2), Inches(4.4))

tx = RAND + Inches(6.7)
tb = BREITE - tx - RAND
tf = textfeld(f, tx, y + Inches(0.4), tb, Inches(2.0))
absatz(tf, "Rabatt", groesse=18, farbe=GRAU, erster=True)
absatz(tf, "65 %", groesse=32, fett=True, farbe=BLAU, abstand_vor=2, zeilenabstand=1.05)
absatz(tf, "Gewicht", groesse=18, farbe=GRAU, abstand_vor=14)
absatz(tf, "23 %", groesse=32, fett=True, farbe=BLAU, abstand_vor=2, zeilenabstand=1.05)
absatz(tf, "alle übrigen 13 Merkmale", groesse=18, farbe=GRAU, abstand_vor=14)
absatz(tf, "unter 5 %", groesse=32, fett=True, farbe=HELLGRAU, abstand_vor=2,
       zeilenabstand=1.05)

tf = textfeld(f, tx, y + Inches(3.6), tb, Inches(0.9))
absatz(tf, "Entweder sicher oder Münzwurf.", groesse=20, fett=True, farbe=ORANGE,
       erster=True, zeilenabstand=1.25)
notizen(f, "7:30 - 8:30  |  60 Sekunden", """Feature Importance zeigt, woran das Modell
seine Entscheidung festmacht: zwei Drittel am Rabatt, ein Viertel am Gewicht, alles
andere zusammen unter 5 Prozent.

Auffaellig: Die Versandart - Flugzeug, Schiff oder Strasse - spielt praktisch keine
Rolle. Fuer einen Logistikdatensatz ist das schwer erklaerbar. Darauf komme ich im
Fazit zurueck.

Der wichtigste Punkt dieser Folie: Das Modell ist entweder voellig sicher oder es raet.
Bei einem Viertel der Bestellungen liegt es zu 100 Prozent richtig. Bei den restlichen
drei Vierteln ist es ein Muenzwurf.

Das erklaert auch den schwachen Recall von vorhin.""")
fusszeile(f, "3 / 4")

# ------------------------------------------------------------------ 4. Fazit
f = leere_folie()
y = kopfzeile(f, "4. Fazit")

box = f.shapes.add_shape(1, RAND, y + Inches(0.1), INHALT_BREITE, Inches(1.4))
box.fill.solid(); box.fill.fore_color.rgb = RGBColor(0xFD, 0xF0, 0xE9)
box.line.color.rgb = ORANGE; box.line.width = Pt(1.5); box.shadow.inherit = False
tf = box.text_frame; tf.word_wrap = True
tf.margin_left = tf.margin_right = Inches(0.35); tf.margin_top = Inches(0.25)
absatz(tf, "Lässt sich eine Verspätung frühzeitig erkennen?", groesse=15, fett=True,
       farbe=ORANGE, erster=True)
absatz(tf, "Teilweise — für ein Viertel der Bestellungen.", groesse=28, fett=True,
       abstand_vor=8, zeilenabstand=1.15)

ky = y + Inches(2.0)
kb, luecke = Inches(3.7), Inches(0.4)
for i, (zahl, text, farbe) in enumerate([
        ("26 %", "zuverlässig vorhersagbar", BLAU),
        ("74 %", "kaum besser als Raten", HELLGRAU),
        ("48 %", "der Verspätungen erkannt", ORANGE)]):
    kachel(f, RAND + i * (kb + luecke), ky, kb, Inches(1.35), zahl, text, farbe)

tf = textfeld(f, RAND, ky + Inches(1.75), INHALT_BREITE, Inches(1.2))
absatz(tf, "Das Modell scheitert nicht am Verfahren, sondern am Datensatz.",
       groesse=21, fett=True, erster=True, zeilenabstand=1.2)
absatz(tf, "Entfernung, Wetter, Zoll, Saison — nicht erfasst.",
       groesse=18, farbe=GRAU, abstand_vor=8)
notizen(f, "8:30 - 9:30  |  60 Sekunden", """Die Ausgangsfrage direkt beantworten:
teilweise. Fuer ein Viertel der Bestellungen ja, fuer den Rest nein.

Ehrlich einordnen: Insgesamt werden nur 48 Prozent der Verspaetungen erkannt. Fuer ein
Fruehwarnsystem reicht das nicht.

Wichtig: Das Modell scheitert nicht am Verfahren, sondern am Datensatz. Was
Verspaetungen wirklich verursacht - Entfernung, Wetter, Zoll, Auslastung - steht
nirgends drin.

Zum Schluss der kritische Hinweis: Eine Regel, die in 2.647 von 2.647 Faellen
ausnahmslos gilt, und eine Versandart ohne jeden Einfluss - das gibt es in echten
Betriebsdaten nicht. Der Datensatz duerfte teilweise kuenstlich erzeugt sein.

Das ist kein Scheitern der Analyse, sondern ein Ergebnis der Analyse.""")
fusszeile(f, "4 / 4")

# ------------------------------------------- 4b. Verbesserungsmoeglichkeiten
f = leere_folie()
y = kopfzeile(f, "4. Verbesserungsmöglichkeiten", "Nach Wirkung sortiert")

eintraege = [
    ("1", "Bessere Daten erheben", "Bestelldatum, Saison, Entfernung, Wetter", ORANGE),
    ("2", "Entscheidungsschwelle senken", "mehr Recall, kostenlos, kein neues Training", BLAU),
    ("3", "Gradient Boosting testen", "XGBoost oder LightGBM", BLAU),
    ("4", "Schwellen automatisch bestimmen", "statt von Hand abgelesen", BLAU),
]

ey = y + Inches(0.35)
for nummer, titel, text, farbe in eintraege:
    kreis = f.shapes.add_shape(9, RAND, ey, Inches(0.6), Inches(0.6))
    kreis.fill.solid(); kreis.fill.fore_color.rgb = farbe
    kreis.line.fill.background(); kreis.shadow.inherit = False
    ktf = kreis.text_frame
    absatz(ktf, nummer, groesse=20, fett=True, farbe=WEISS, erster=True, zeilenabstand=1.0)
    ktf.paragraphs[0].alignment = PP_ALIGN.CENTER

    tf = textfeld(f, RAND + Inches(0.95), ey - Inches(0.08),
                  INHALT_BREITE - Inches(0.95), Inches(1.0))
    absatz(tf, titel, groesse=21, fett=True, erster=True, zeilenabstand=1.1)
    absatz(tf, text, groesse=17, farbe=GRAU, abstand_vor=4, zeilenabstand=1.2)
    ey += Inches(1.15)

notizen(f, "9:30 - 10:00  |  30 Sekunden", """Nur die ersten beiden Punkte ansprechen,
die anderen stehen zum Nachlesen da.

Punkt 1 ist der groesste Hebel: bessere Daten. Schon das Bestelldatum allein - also
Saison und Wochentag - wuerde vermutlich mehr bringen als alle bisherigen Merkmale
zusammen.

Punkt 2 ist der schnellste Gewinn: Die Entscheidungsschwelle von 50 Prozent senken.
Das Modell warnt dadurch haeufiger, findet mehr Verspaetungen und kostet nichts -
kein neues Training noetig. Genau das Recall-Problem von vorhin.

Abschluss: "Vielen Dank - gerne Fragen."

PUFFER: Falls die Zeit knapp wird, diese Folie in 15 Sekunden abhandeln.""")
fusszeile(f, "4 / 4")

ziel = "praesentation/Lieferprognose_Praesentation.pptx"
prs.save(ziel)
print("Gespeichert:", ziel, "|", len(prs.slides._sldIdLst), "Folien")
