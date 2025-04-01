import { initializeClient } from "./functions";
import { setStorageItem, removeStorageItem } from "./client";

// Catch any unexpected global errors or promise rejections
window.addEventListener('unhandledrejection', event => {
  console.error('Unhandled promise rejection:', event.reason);
});

window.onerror = (msg, url, lineNo, columnNo, error) => {
  console.error('Global JS error:', msg, error);
};

Office.onReady(() => {
  document.addEventListener('DOMContentLoaded', () => {
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

    const safeForm = form as HTMLFormElement;
    const safeLogoutButton = logoutButton as HTMLButtonElement;
    const safeStatus = status as HTMLElement;

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
        setStorageItem('ArchitectApiKey', apiKey);
        setStorageItem('ArchitectApiSecret', apiSecret);

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
        removeStorageItem('ArchitectApiKey');
        removeStorageItem('ArchitectApiSecret');
        safeForm.reset();
        setStatus('Logged out!');
      } catch (err) {
        console.error("Logout failed:", err);
        setStatus(`Error: ${(err as Error).message}`);
      }
    }

    function setStatus(message: string) {
      safeStatus.textContent = message;
    }
  });
});
