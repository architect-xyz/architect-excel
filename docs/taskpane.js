import { initializeClient } from './architect-excel.js';

document.getElementById('api-form').addEventListener('submit', async (e) => {
    e.preventDefault();
    const apiKey = document.getElementById('apiKey').value;
    const apiSecret = document.getElementById('apiSecret').value;
  
    try {
      // Save credentials (secure storage recommended)
      Office.context.document.settings.set('apiKey', apiKey);
      Office.context.document.settings.set('apiSecret', apiSecret);
      Office.context.document.settings.saveAsync();
      document.getElementById('status').textContent = 'Credentials saved!';

      initializeClient()
    } catch (err) {
      document.getElementById('status').textContent = `Error: ${err.message}`;
    }

  });