const SETTINGS_KEY = "everyrow_settings";

interface Settings {
  apiKey?: string;
}

function getSettings(): Settings {
  try {
    // Use Office roaming settings for cross-device persistence
    if (Office.context?.roamingSettings) {
      const data = Office.context.roamingSettings.get(SETTINGS_KEY);
      return data ? JSON.parse(data) : {};
    }
    // Fallback to localStorage for development
    const data = localStorage.getItem(SETTINGS_KEY);
    return data ? JSON.parse(data) : {};
  } catch {
    return {};
  }
}

function saveSettings(settings: Settings): void {
  try {
    const data = JSON.stringify(settings);
    if (Office.context?.roamingSettings) {
      Office.context.roamingSettings.set(SETTINGS_KEY, data);
      Office.context.roamingSettings.saveAsync();
    } else {
      localStorage.setItem(SETTINGS_KEY, data);
    }
  } catch (error) {
    console.error("Failed to save settings:", error);
    throw error;
  }
}

export function getApiKey(): string | null {
  const settings = getSettings();
  return settings.apiKey || null;
}

export function saveApiKey(apiKey: string): void {
  const settings = getSettings();
  settings.apiKey = apiKey;
  saveSettings(settings);
}

export function clearApiKey(): void {
  const settings = getSettings();
  delete settings.apiKey;
  saveSettings(settings);
}

export function validateApiKeyFormat(apiKey: string): boolean {
  return apiKey.startsWith("sk-cho-");
}
