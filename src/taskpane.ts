import { initializeClient } from "./functions";
import { setStorageItem, removeStorageItem } from "./client";

Office.onReady(() => {
  const form = document.getElementById('api-form') as HTMLFormElement;
  const logoutButton = document.getElementById('logout-button') as HTMLButtonElement;
  const status = document.getElementById('status')!;

  form.addEventListener('submit', handleFormSubmit);
  logoutButton?.addEventListener('click', handleLogout);

  function cleanField(field: FormDataEntryValue | null): string {
    return (field as string)?.trim() || '';
  }

  async function handleFormSubmit(e: Event) {
    e.preventDefault();

    const formData = new FormData(form);
    const apiKey = cleanField(formData.get('apiKey'));
    const apiSecret = cleanField(formData.get('apiSecret'));

    if (!apiKey || !apiSecret) {
      setStatus('API Key and Secret are required.');
      return;
    }

    try {
      setStorageItem('ArchitectApiKey', apiKey);
      setStorageItem('ArchitectApiSecret', apiSecret);

      const success = await initializeClient();
      setStatus(success
        ? 'Credentials saved! Client initialized!'
        : 'Credentials saved! However, Client was NOT successfully initialized!');
    } catch (err) {
      setStatus(`Error: ${(err as Error).message}`);
    }
  }

  async function handleLogout(e: Event) {
    e.preventDefault();

    try {
      removeStorageItem('ArchitectApiKey');
      removeStorageItem('ArchitectApiSecret');
      form.reset();
      setStatus('Logged out!');
    } catch (err) {
      setStatus(`Error: ${(err as Error).message}`);
    }
  }

  function setStatus(message: string) {
    status.textContent = message;
  }
});
