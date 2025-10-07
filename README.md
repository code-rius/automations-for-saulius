# automations-for-saulius

## Setup Instructions

### 1. Create `.env` file

Copy `example.env` to `.env` and fill in your cookie value:

```sh
copy example.env .env
```
Or manually create a file named `.env` in the project folder with this content:

```
RC_COOKIE=your_cookie_value
```

### 2. Install Python libraries

Install required libraries using pip:

```sh
pip install -r requirements.txt
```

If `requirements.txt` does not exist, install manually:

```sh
pip install requests python-dotenv beautifulsoup4 pdfplumber
```

### 3. Usage

Run the scripts as needed, for example:

```sh
python address-extractor.py
python pdfreader.py
```

---

**Note:**  
- `.env` is used to securely store your cookie and other environment variables.
- Make sure `.env` is not tracked by git (see `.gitignore`).

TODO:
1. Sutvarkyti output.txt (done)
2. Padaryti, kad veiktu su keliomis elektrinemis is .env (done)
3. Agreguoti kelias keletrine i viena etapo folderi (done)
4. Padaryti, kad sugeneruotu dokumentus (+-done)
5. Paziureti imoniu adreso istraukima is registru centro