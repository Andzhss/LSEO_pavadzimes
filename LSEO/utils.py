from num2words import num2words
import requests
from bs4 import BeautifulSoup
import re

def money_to_words_lv(amount):
    """
    Konvertē summu uz vārdiem latviešu valodā.
    Piemērs: 4505.00 -> "Četri tūkstoši pieci simti pieci eiro 00 centi"
    """
    try:
        euros = int(amount)
        cents = int(round((amount - euros) * 100))
        
        words = num2words(euros, lang='lv')
        
        # Pirmā burta lielais sākums
        words = words.capitalize()
        
        return f"{words} eiro {cents:02d} centi"
    except Exception as e:
        return f"Kļūda aprēķinā: {e}"

def scrape_lursoft(url):
    """
    Nolasa uzņēmuma / iestādes datus no dažādiem Lursoft lapu veidiem:
      - https://company.lursoft.lv/en/<slug>/<reg-no>
      - https://company.lursoft.lv/lv/<slug>/<reg-no>
      - https://iestades.lursoft.lv/<slug>/<reg-no>

    Atgriež vārdnīcu ar atslēgām: name, reg_no, vat_no, address (vai None).
    """
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }
        response = requests.get(url, headers=headers, timeout=15)
        response.raise_for_status()

        soup = BeautifulSoup(response.text, 'html.parser')
        data = {}

        # --- Reģistrācijas numurs no URL (uzticamākais avots) ---
        m = re.search(r'/(\d{6,14})(?:[/?#]|$)', url)
        if m:
            data['reg_no'] = m.group(1)

        # --- Palīgfunkcija: vērtība tabulas rindā, kur pirmā šūna atbilst etiķetei ---
        def field_value(*labels):
            wanted = {str(l).strip().lower() for l in labels}
            for tr in soup.find_all('tr'):
                cells = tr.find_all(['th', 'td'])
                if len(cells) >= 2:
                    key = cells[0].get_text(' ', strip=True).lower()
                    if key in wanted:
                        return cells[1].get_text(' ', strip=True)
            return None

        # --- Nosaukums ---
        h1 = soup.find('h1')
        name = h1.get_text(strip=True) if h1 else None
        if not name:
            name = field_value('Nosaukums', 'Name', 'Institūcijas nosaukums')
            if name:
                name = re.split(r'Iepriekšējie nosaukumi|Previous names', name, flags=re.I)[0].strip()
        if not name and soup.title:
            name = soup.title.get_text(strip=True).split(',')[0].strip()
        if name:
            data['name'] = name

        # --- Reģ. Nr. no lapas (rezerves variants, ja URL nav) ---
        if not data.get('reg_no'):
            reg_val = field_value(
                'Reģistrācijas numurs, datums',
                'Registration number, date',
                'Reģistrācijas numurs'
            )
            if reg_val:
                rm = re.search(r'\d{6,14}', reg_val)
                if rm:
                    data['reg_no'] = rm.group(0)

        # --- PVN numurs ---
        whole_text = soup.get_text(' ', strip=True)
        mvat = re.search(r'LV\d{6,14}', whole_text)
        if mvat:
            data['vat_no'] = mvat.group(0)
        elif data.get('reg_no'):
            data['vat_no'] = 'LV' + data['reg_no']

        # --- Adrese ---
        # 1) Tabulas rinda "Juridiskā adrese" / "Legal address"
        addr = field_value('Juridiskā adrese', 'Legal address')
        # 2) Ja vērtība netika atrasta, meklējam "Juridiskā adrese:" tekstu
        if not addr:
            addr_node = soup.find(string=re.compile(r'^(Juridiskā adrese|Legal address)\s*:', re.I))
            if addr_node:
                addr = addr_node.parent.get_text(' ', strip=True)
        if addr:
            clean = re.sub(r'^(Juridiskā adrese|Legal address)\s*:?\s*', '', addr, flags=re.I)
            clean = re.split(
                r'Iepriekšējās adreses|Previous addresses|Adresē reģistrēti|Other registered|Pasta adrese',
                clean, flags=re.I
            )[0]
            clean = clean.strip().rstrip(',').strip()
            if clean and 'Uzņēmumi' not in clean and 'Institūcijas' not in clean:
                data['address'] = clean

        # Atgriežam tikai tad, ja vismaz nosaukums ir atrasts
        return data if data.get('name') else None
    except Exception as e:
        print(f"Scraping error: {e}")
        return None
