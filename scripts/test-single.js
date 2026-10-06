import axios from 'axios';

async function checkFixtureStructure() {
    try {
        const res = await axios.get('https://sportscore.com/api/v1/fixtures/?sport=football&date=2026-09-11', {
            headers: { 'User-Agent': 'Mozilla/5.0' },
            timeout: 10000
        });

        const list = res.data?.events || res.data?.fixtures || res.data?.matches || res.data?.data || [];
        console.log(`Totale eventi ricevuti: ${list.length}\n`);

        if (list.length > 0) {
            console.log('--- ESEMPIO OGGETTO EVENTO (CHIAVI DISPONIBILI) ---');
            console.log(Object.keys(list[0]));
            console.log('\n--- CONTENUTO PRIMO EVENTO ---');
            console.log(JSON.stringify(list[0], null, 2));
        } else {
            console.log('Struttura radice:', Object.keys(res.data || {}));
        }
    } catch (err) {
        console.error('Errore chiamata:', err.message);
    }
}

checkFixtureStructure();