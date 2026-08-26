# Veille IP DLA

Dashboard Netlify e automazione GitHub Actions per generare una veille settimanale di proprietà intellettuale da fonti pubbliche verificate.

## Attivazione

1. Collegare questo repository a Netlify con publish directory `.`.
2. Aggiungere in GitHub Actions il secret `OPENAI_API_KEY`.
3. Avviare manualmente il workflow `Generate weekly veille` per il primo test.

Le fonti riservate, i paywall e le credenziali del cabinet sono esclusi.
