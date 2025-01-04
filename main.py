import random
from aiolimiter import AsyncLimiter
from playwright.async_api import async_playwright
import pandas as pd
from tqdm import tqdm
import re

limiter = AsyncLimiter(max_rate=10, time_period=1)  # 1 saniyede 10 istek
USER_AGENTS = [
    # Windows 
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:89.0) Gecko/20100101 Firefox/89.0",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Edge/91.0.864.59 Safari/537.36",

    # macOS 
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7; rv:89.0) Gecko/20100101 Firefox/89.0",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 11_0_0) AppleWebKit/537.36 (KHTML, like Gecko) Safari/537.36",

    # Linux 
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
    "Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:89.0) Gecko/20100101 Firefox/89.0",
    "Mozilla/5.0 (X11; Linux i686) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.85 Safari/537.36",

    # Mobil 
    "Mozilla/5.0 (iPhone; CPU iPhone OS 14_6 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/14.0 Mobile/15E148 Safari/604.1",
    "Mozilla/5.0 (iPad; CPU OS 14_6 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/14.0 Mobile/15E148 Safari/604.1",
    "Mozilla/5.0 (Linux; Android 11; Pixel 5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Mobile Safari/537.36",
    "Mozilla/5.0 (Linux; Android 10; SM-G960F) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.85 Mobile Safari/537.36",

    # Eski 
    "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.1; Trident/4.0)",
    "Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:78.0) Gecko/20100101 Firefox/78.0",

    # Diğer
    "Mozilla/5.0 (PlayStation 4 3.11) AppleWebKit/537.73 (KHTML, like Gecko)",
    "Mozilla/5.0 (Linux; U; Android 6.0; en-US; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/91.0.4472.124 Mobile Safari/537.36",
    "Mozilla/5.0 (Linux; U; Android 9; en-US; SM-J600F Build/PPR1.180610.011) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.85 Mobile Safari/537.36"
]
async def fetch_website(company_name):
    """Playwright ile Google üzerinden web sitesi arar."""
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        user_agent = random.choice(USER_AGENTS)  # Rastgele bir User-Agent seç
        context = await browser.new_context(user_agent=user_agent)  # User-Agent ayarla
        page = await context.new_page()
        try:
            search_query = f"https://www.google.com/search?q={company_name}+website"
            await page.goto(search_query)
            # İlk Google sonucunu al
            result = await page.locator(".tF2Cxc a").first.get_attribute("href")
            await browser.close()
            return result
        except Exception as e:
            print(f"Hata: {company_name} - {e}")
            await browser.close()
            return None

async def fetch_contact_info(website):
    """Web sitesinden telefon ve e-posta bilgilerini çeker."""
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        user_agent = random.choice(USER_AGENTS)  # Rastgele bir User-Agent seç
        context = await browser.new_context(user_agent=user_agent)  # User-Agent ayarla
        page = await context.new_page()
        try:
            await page.goto(website, timeout=15000)
            await page.wait_for_load_state("networkidle")
            content = await page.content()

            # E-posta adreslerini bul
            emails = set(re.findall(r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+", content))
            # Telefon numaralarını bul
            phones = set(re.findall(r"(?:(?:\+|00)[\d\s()-]{7,}|(?:\b\d{2,4}\)?[\s.-]?\d{2,4}[\s.-]?\d{2,4}\b))", content))
            valid_phones = {phone for phone in phones if len(re.sub(r'\D', '', phone)) >= 10}

            await browser.close()
            return ", ".join(emails) if emails else None, ", ".join(valid_phones) if valid_phones else None
        except Exception as e:
            print(f"Hata: {website} - {e}")
            await browser.close()
            return None, None

async def main():
    # Excel dosyasını oku
    file_path = "CompanyList.xlsx"  # Excel dosyasının yolu
    output_file = "results.xlsx"  # Çıkış dosyası

    try:
        data = pd.read_excel(file_path, header=1)
        company_names = data['Firma İsmi'].dropna().tolist()[:50]  # İlk 50 şirketi seç
    except Exception as e:
        print(f"Hata: {e}")
        return

    print(f"{len(company_names)} şirket ismi bulundu.")

    results = []

    # Şirket web sitelerini bul ve iletişim bilgilerini topla
    for company in tqdm(company_names, desc="Şirket bilgileri işleniyor"):
        website = await fetch_website(company)
        if website:
            email, phone = await fetch_contact_info(website)
            results.append({
                "Firma İsmi": company,
                "Web Sitesi": website,
                "Telefon": phone,
                "E-posta": email
            })
        else:
            results.append({
                "Firma İsmi": company,
                "Web Sitesi": None,
                "Telefon": None,
                "E-posta": None
            })

    # Sonuçları kaydet
    pd.DataFrame(results).to_excel(output_file, index=False)
    print(f"Sonuçlar kaydedildi: {output_file}")

if __name__ == "__main__":
    import asyncio
    asyncio.run(main())