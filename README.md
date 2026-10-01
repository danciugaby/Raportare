# Analiza evaluării activității didactice de către studenți

Script Python care citește exportul platformei de evaluare UNITBV și produce un raport HTML, plus fișiere CSV pentru Excel.

Sursele citite:

- fișele individuale PDF (formularul F05.1-PS8.2-01/ed.4, rev.2);
- rapoartele pe program (`...(cod).docx`);
- tabelele de sinteză (`...(cod)_SINTEZA.docx`);
- opțional, raportul de sinteză al facultății (`.docx`).

## Instalare (o singură dată)

Ai nevoie de Python 3.9 sau mai nou.

```bash
cd analiza_evaluare
python -m venv .venv
# Windows:      .venv\Scripts\activate
# macOS/Linux:  source .venv/bin/activate
pip install -r requirements.txt
```

## Rulare

Direct din arhiva exportată:

```bash
python analiza_evaluare.py --zip IESC.zip --raport-facultate "Facultatea_de_Inginerie_Electrica_si_Stiinta_Calculatoarelor10_CU_FRECVENTA.docx"
```

Sau dintr-un folder deja dezarhivat:

```bash
python analiza_evaluare.py --dir IESC --raport-facultate raport_facultate.docx --out rezultate
```

Citirea celor ~480 de fișe PDF durează 2–3 minute. Dacă ai un procesor cu mai multe nuclee, rulează cu `--procese 8` ca să meargă mai repede.

## Ce obții în folderul `rezultate/`

| Fișier | Conținut |
|---|---|
| `raport_evaluare.html` | Raportul vizual, care se deschide în orice browser: participare, niveluri pe dimensiuni, programe, departamente, cazuri de analizat, teme din comentarii, verificări, măsuri și tabelul cu toate fișele (cu filtre și sortare). |
| `fise.csv` | O linie pentru fiecare poziție cadru didactic × disciplină × curs/aplicații × program: respondenți, scor D1–D7, % nivel scăzut la D6, D8, numele fișierului. |
| `comentarii.csv` | Toate comentariile, cu rubrica, temele etichetate automat și o coloană goală `tema_manuala`. |
| `programe.csv`, `dimensiuni.csv` | Agregatele pe program și pe dimensiune (S/M/R, curs vs aplicații). |
| `verificari.txt` | Reconcilierea cifrelor între surse și problemele găsite (câmpuri goale, indicatori calculați greșit, fișe lipsă, codificare). |
| `date_complete.json` | Tot ce s-a extras, pentru alte analize. |
| `_arhiva/` | Fișierele dezarhivate (doar la `--zip`), cu numele reparate. |

Fișierele CSV folosesc `;` ca separator și UTF-8 cu BOM, așa că se deschid corect în Excel cu setări românești.

## Partea scrisă de mână: `observatii.html`

Scriptul calculează automat toate cifrele, graficele și verificările. Nu poate însă să judece singur comentariile. Semnalările privind conduita, cazurile în care s-a evaluat altă persoană, practicile bune și măsurile propuse vin din fișierul `observatii.html`.

Versiunea inclusă conține observațiile pentru 2025–2026, semestrul 2, scrise după citirea integrală a comentariilor. Pentru alt semestru:

1. Rulează scriptul o dată.
2. Citește `comentarii.csv`.
3. Rescrie blocurile din `observatii.html` (REZUMAT, SEMNALARI, ATRIBUIRE, PRACTICI, MASURI). Marcajele `<!-- NUME -->` … `<!-- /NUME -->` trebuie păstrate.
4. Rulează din nou scriptul.

Pentru un raport fără aceste secțiuni, folosește `--observatii fisier_gol.html`.

## Teme din comentarii: automat sau manual

Implicit, temele sunt etichetate automat după cuvinte cheie. Rezultatul e orientativ: prinde bine aprecierile, dar mai puțin bine criticile. Pentru cifre de folosit în raportul oficial:

1. Deschide `rezultate/comentarii.csv` în Excel.
2. Completează coloana `tema_manuala`. Pune numele temelor, iar dacă sunt mai multe le separi cu `|`.
3. Salvează fișierul ca CSV UTF-8, cu separator `;`.
4. Rulează din nou: `python analiza_evaluare.py --zip IESC.zip --raport-facultate ... --teme-manuale rezultate/comentarii.csv`

## Opțiuni utile

| Opțiune | Implicit | Ce face |
|---|---|---|
| `--prag-n` | 3 | numărul minim de respondenți pentru lista „Cazuri de analizat” |
| `--prag-scor` | 3.5 | o fișă intră în listă dacă scorul D1–D7 e sub acest prag… |
| `--prag-s` | 50 | …sau dacă cel puțin acest % din aprecierile globale (D6) sunt la nivel scăzut |
| `--an` | citit din sinteze | textul afișat pentru an/semestru |
| `--out` | `rezultate` | folderul de ieșire |

## Cum funcționează (pe scurt)

Codul din `analiza_evaluare.py` e împărțit în secțiuni numerotate:

1. **Arhiva.** Dezarhivează și repară numele de fișiere. Arhiva are nume în UTF-8 fără marcaj, iar unele au literele Ă/Ț/Î stocate greșit.
2. **Fișele PDF.** Citește antetul, numărul de respondenți, scorurile D1–D7 și D8, nivelul bifat pe fiecare dimensiune și comentariile. Folosește poziția textului în pagină (pdfplumber).
3. **Rapoartele docx.** Citește tabelele 1.1–1.3 și 2–3 din rapoartele pe program și numărul de respondenți pe nivel (S/M/R) din tabelele de sinteză. Coloanele goale din aceste tabele sunt deduse din antet.
4. **Îmbinarea.** Leagă fiecare rând din sinteză de fișa PDF după program, tip, cadru didactic și disciplină, comparate fără diacritice. Rândurile fără fișă sunt raportate ca fișe lipsă.
5. **Analize.** Medii ponderate, acoperire, distribuții, niveluri, teme.
6. **Verificări.** Reconcilieri între surse și probleme de date.
7. **Exporturi și raport HTML** (din `sablon_raport.html`).

Prescurtările programelor sunt în `PROGRAM_ALIAS`, la începutul scriptului. Un program nou primește automat o prescurtare din inițiale. Dacă vrei alta, o adaugi acolo.

**Limită:** citirea fișelor PDF depinde de macheta actuală a formularului. Dacă platforma schimbă formularul, poate fi nevoie să ajustezi coordonatele din funcția `citeste_fisa`. Rândul „Reconcilieri” din `verificari.txt` arată imediat dacă cifrele nu se mai potrivesc.

## Confidențialitate

Rezultatele conțin evaluări nominale ale cadrelor didactice și comentarii ale studenților. Păstrează folderele `rezultate/` și `_arhiva/` doar local sau pe un spațiu cu acces restrâns. Nu le pune într-un depozit public.
