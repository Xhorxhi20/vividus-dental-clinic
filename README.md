# 🦷 Vividus Dental Clinic

Sistem desktop për menaxhimin e një klinike dentare, i ndërtuar me Python dhe CustomTkinter.

## Funksionalitetet

- **Menaxhim pacientësh** — regjistrim, modifikim, kërkim në kohë reale
- **Planifikim terminesh** — me validim të datës, orës dhe çmimit
- **Historik trajtimesh** — të gjitha vizitat e një pacienti me totalin e paguar
- **Dashboard** — statistika për pacientët, terminet dhe të ardhurat
- **Filtrim sipas statusit** — Planifikuar / Përfunduar / Anulluar
- **Eksport në Excel** — raporte të plota me një klik
- **Backup automatik** — ruajtje e databazës në çdo hapje dhe mbyllje
- **Kujtesë ditore** — njoftim për terminet e ditës kur niset programi

## Teknologjitë

`Python` · `CustomTkinter` · `SQLite` · `openpyxl`

## Instalimi

```bash
pip install customtkinter openpyxl
python vividus.py
```

## Ndërtimi si aplikacion .exe

```bash
pyinstaller --onefile --windowed --icon=Logo.ico --name "Vividus Clinic" vividus.py
```

## Ruajtja e të dhënave

Databaza krijohet automatikisht te `AppData/Local/VividusClinic/` herën e parë që niset programi. Ajo nuk përfshihet në këtë repository për arsye privatësie.

## Autori

Zhvilluar nga [Xhorxhi20](https://github.com/Xhorxhi20)