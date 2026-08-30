# translate-to-spanish-script-for-google-sheets

WIP: Will need to make sure the script is wary of rate limits, by doing runs in batches.


App script to translate a Sheet in the active spreadsheet in Google Sheets.

This is good for small data sets, but does not scale well as it duplicates the target sheet to make a translated one.


To translate to a different language, change 'es' in this line: 

const translated = LanguageApp.translate(cellValue, '', 'es');

to another language code:

    'fr' for French

    'de' for German

    'it' for Italian

    'pt' for Portuguese

    'zh' for Chinese

    etc.
