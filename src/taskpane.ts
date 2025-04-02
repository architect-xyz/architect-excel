import { initializeClient, remakeClient } from "./functions";
import { config, removeStorageItem, setStorageItem } from "./client";

Office.onReady(() => {
  const form = document.getElementById('api-form') as HTMLFormElement | null;
  const logoutButton = document.getElementById('logout-button') as HTMLButtonElement | null;
  const status = document.getElementById('status');

  if (!form || !logoutButton || !status) {
    console.error("Missing expected DOM elements:", {
      formExists: !!form,
      logoutButtonExists: !!logoutButton,
      statusExists: !!status
    });
    return;
  }

  const safeForm = form;
  const safeLogoutButton = logoutButton;
  const safeStatus = status;

  safeForm.addEventListener('submit', handleFormSubmit);
  safeLogoutButton.addEventListener('click', handleLogout);

  function cleanField(field: FormDataEntryValue | null): string {
    return (field as string)?.trim() || '';
  }

  async function handleFormSubmit(e: Event) {
    e.preventDefault();

    const formData = new FormData(safeForm);
    const apiKey = cleanField(formData.get('apiKey'));
    const apiSecret = cleanField(formData.get('apiSecret'));

    if (!apiKey || !apiSecret) {
      setStatus('API Key and Secret are required.');
      return;
    }
    try {
      await setStorageItem('ArchitectApiKey', apiKey);
      await setStorageItem('ArchitectApiSecret', apiSecret);

      const success = await initializeClient();
      setStatus(success
        ? 'Credentials saved! Client initialized!'
        : 'Credentials saved, but client failed to initialize.');
    } catch (err) {
      console.error("Initialization failed:", err);
      setStatus(`Error: ${(err as Error).message}`);
    }
  }

  async function handleLogout(e: Event) {
    e.preventDefault();

    try {
      await removeStorageItem('ArchitectApiKey');
      await removeStorageItem('ArchitectApiSecret');
      safeForm.reset();
      remakeClient("", "")
      setStatus('Logged out! The Add-in will no longer have access to your credentials on the next restart.');
    } catch (err) {
      console.error("Logout failed:", err);
      setStatus(`Error: ${(err as Error).message}`);
    }
  }

  function setStatus(message: string) {
    safeStatus.textContent = message;
  }
});
