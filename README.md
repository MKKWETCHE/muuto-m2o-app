# ⚙️ Muuto Made-to-Order (MTO) Master Data Tool

Dette projekt er en avanceret **Streamlit-applikation** designet til **Channel Marketing** og brug af **kunder** for at forenkle valget af Muuto Made-to-Order (MTO) produkter og generere den nødvendige masterdata. Appen kører live på **https://muuto-m2o-app.streamlit.app/**.

---

## 💡 Applikationens Formål

Applikationen fungerer som en **konfigurator** for at sammensætte komplekse MTO-produkter (produkt, polstring, farve og basefarve) og eksportere en komplet masterdata-fil, herunder dynamiske priser baseret på valgt valuta.

### Arbejdsgang (Steps)

1.  **Vælg Valuta (Step 1):** Bestemmer, hvilke markeder (EU eller UK/IE) og produkter der er tilgængelige.
2.  **Vælg Produktkombinationer (Step 2):** Brugeren navigerer via Produktfamilie og vælger produkter i en matrix baseret på Polstringstype og Farve.
3.  **Specificér Basefarver (Step 2a):** For produkter, der kræver et valg af basefarve, specificeres denne enten pr. produkt eller samlet pr. familie.
4.  **Gennemse Valg (Step 3):** Viser en opsummeret liste over de valgte SKU'er, hvorfra individuelle elementer kan fjernes.
5.  **Generer Fil (Step 4):** Opretter og downloader en Excel-fil med beriget masterdata og priser.

---

## 🛠️ Opsætning og Filer

For at køre eller vedligeholde appen lokalt, kræves følgende filstruktur i samme mappe som scriptet:

### 1. Nødvendige Datafiler

| Filnavn | Formål | Vigtig Sheet | Nøglekolonner (Eksempler) |
| :--- | :--- | :--- | :--- |
| `raw-data.xlsx` | **Rå Produktdata** (Alle mulige kombinationer, SKU'er, Basefarver, Billed-URL'er). | `APP` | `Item No`, `Article No`, `Product Type`, `Upholstery Type`, `Base Color`, `Market`. |
| `price-matrix_EUROPE.xlsx` | Priser for **EU-valutaer** (EURO, DKK, SEK, NOK, PLN, AUD, DACH - EURO). | `Price matrix wholesale`, `Price matrix retail` | Valutakolonner (`EURO`, `DKK`, osv.) og Artikelnummer. |
| `price-matrix_GBP-IE.xlsx` | Priser for **UK/IE-valutaer** (GBP, IE - EUR). | `Price matrix wholesale`, `Price matrix retail` | Valutakolonner (`GBP`, `IE - EUR`) og Artikelnummer. |
| `Masterdata-output-template.xlsx`| Definerer rækkefølgen af kolonner i den endelige output-fil. | Standard | Indeholder *alle* ønskede kolonner, inkl. `Wholesale price` og `Retail price` (som erstattes dynamisk). |
| `muuto_logo.png` | Logo-fil til visning i UI. | N/A | |

### 2. Python-Biblioteker

Installer nødvendige afhængigheder for lokal kørsel:

```bash
pip install streamlit pandas openpyxl xlsxwriter
````

-----

## ⚙️ Kernen i Logikken

### A. Datafiltrering (Step 1)

Valget af valuta styrer, hvilke produkter der vises, baseret på `Market`-kolonnen i `raw-data.xlsx`:

  * Hvis en **EUROPE-valuta** vælges: Viser produkter, hvor `Market` **ikke** er `"UK"`.
  * Hvis en **UK/IE-valuta** vælges: Viser produkter, hvor `Market` **ikke** er `"EU"`.

### B. Produktvisning (`construct_product_display_name`)

En **`Product Display Name`** oprettes dynamisk til visning i matrixen ved at kombinere relevante kolonner, såsom `Product Type`, `Product Model` og (for sofaer) `Sofa Direction`.

### C. Matrixlogik (Step 2)

Hver celle i matrixen repræsenterer en **generisk varekombination** (f.eks., "Outline Sofa - Læder - Cognac").

  * **`handle_matrix_cb_toggle`:** Callback-funktion, der lagrer den valgte **generiske kombination** i `st.session_state.matrix_selected_generic_items`.
  * **Basefarve-krav:** Når en generisk vare vælges, identificeres det, om den kræver et efterfølgende valg af basefarve (hvis der er flere end én unik `Base Color` for kombinationen).

### D. Basefarve Håndtering (Step 2a)

Dette trin løser SKU'er med variable basefarver:

  * Produkter, der kræver basefarve-valg, grupperes efter `Product Family`.
  * Brugeren kan vælge en eller flere basefarver enten **på familieniveau** (anvendes på alle gældende produkter) eller **individuelt** via multiselects pr. valgt vare.
  * Valgte basefarver gemmes i `st.session_state.user_chosen_base_colors_for_items`.

### E. Output Generering (Step 4)

1.  **Finalisering af SKU'er:** Listen `st.session_state.final_items_for_download` opbygges ved at kombinere de generiske valg (fra Step 2) med de valgte basefarver (fra Step 2a) for at finde de **specifikke `Item No`** og **`Article No`** fra `raw-data.xlsx`.
2.  **Prisopslag:** For hver finaliseret SKU (`Article No`) hentes den korrekte `Wholesale price` og `Retail price` fra den relevante pris-matrix (`EUROPE` eller `GBP-IE`) ved hjælp af den valgte valuta-kolonne.
3.  **Filstruktur:** Outputtet opbygges som et DataFrame med kolonner defineret af `Masterdata-output-template.xlsx`, og de dynamiske pris-kolonner navngives (f.eks., `"Wholesale price (EURO)"`).

-----

## 🎨 UI og Styling

Appen bruger **Streamlit's `st.columns`** til at skabe den komplekse matrix-lignende UI-struktur (polstringstype, farveprøver, farvenumre og produkt-checkboxes). Der er implementeret omfattende **CSS-styling** (via `st.markdown("<style>...</style>")`) for at matche Muuto's branding (farver, skrifttyper, knap-udseende) og sikre, at matrix-elementer og checkboxes er korrekt justeret i et kompakt, "wide" layout.

```
```
