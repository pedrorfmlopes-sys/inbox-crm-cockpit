import dotenv from 'dotenv';
import path from 'path';
// Node 18+ has native fetch

dotenv.config({ path: path.resolve('server/.env') });

async function listModels() {
    const apiKey = process.env.GEMINI_API_KEY;
    console.log('Using Key:', apiKey?.slice(0, 10) + '...');

    const url = `https://generativelanguage.googleapis.com/v1/models?key=${apiKey}`;

    try {
        const res = await fetch(url);
        const data = await res.json();

        if (!res.ok) {
            console.error('Error fetching models:', data.error?.message || res.status);
            return;
        }

        console.log('Available Models:');
        data.models?.forEach(m => {
            console.log(` - ${m.name} (supports: ${m.supportedGenerationMethods.join(', ')})`);
        });
    } catch (e) {
        console.error('Fetch Error:', e.message);
    }
}

listModels();
