document.addEventListener('DOMContentLoaded', () => {
  const form = document.getElementById('api-form');
  form?.addEventListener('submit', async (e) => {
    e.preventDefault();

    const apiKey = (document.getElementById('apiKey') as HTMLInputElement)?.value.trim();
    const apiSecret = (document.getElementById('apiSecret') as HTMLInputElement)?.value.trim();
    const status = document.getElementById('status');

    if (!apiKey || !apiSecret) {
      status!.textContent = 'API Key and Secret are required.';
      return;
    }

    try {
      localStorage.setItem('ArchitectApiKey', apiKey);
      localStorage.setItem('ArchitectApiSecret', apiSecret);
      status!.textContent = 'Credentials saved!';
    } catch (err) {
      status!.textContent = `Error: ${(err as Error).message}`;
    }
  });
});
