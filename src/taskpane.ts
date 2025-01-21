import { initializeClient, setStorageItem } from "./functions";

Office.onReady(() => {
  const form = document.getElementById('api-form') as HTMLFormElement;

  function cleanField(field: FormDataEntryValue | null)  : string {
    return ((field as string) || '').trim();
  }

  form.addEventListener('submit', async (e) => {
    e.preventDefault();
    const formData = new FormData(form);
    const apiKey = cleanField(formData.get('apiKey'));
    const apiSecret = cleanField(formData.get('apiSecret'));

    const status = document.getElementById('status')!;

    if (!apiKey || !apiSecret) {
      status.textContent = 'API Key and Secret are required.';
      return;
    }

    try {
      setStorageItem('ArchitectApiKey', apiKey);
      setStorageItem('ArchitectApiSecret', apiSecret);
      status.textContent = 'Credentials saved!';
      initializeClient();
      status.textContent = 'Credentials saved! Client initialized!';
    } catch (err) {
      status.textContent = `Error: ${(err as Error).message}`;
    }
  });
});
