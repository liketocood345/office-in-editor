import { execFileSync } from "node:child_process";

export interface OfficeTabAppearance {
  background: string;
  foreground: string;
  theme: string;
}

const LIGHT: OfficeTabAppearance = { background: "#FFFFFF", foreground: "#201F1E", theme: "white" };
const DARK_GRAY: OfficeTabAppearance = { background: "#535353", foreground: "#FFFFFF", theme: "dark-gray" };
const BLACK: OfficeTabAppearance = { background: "#191919", foreground: "#FFFFFF", theme: "black" };
const DARK: OfficeTabAppearance = { background: "#1F1F1F", foreground: "#FFFFFF", theme: "dark" };

/** Microsoft Office UI Theme DWORD → Word-like canvas colors. */
const BY_ID: Record<number, OfficeTabAppearance> = {
  0: { ...LIGHT, theme: "colorful" },
  1: DARK_GRAY,
  2: BLACK,
  3: LIGHT,
  5: DARK,
};

function registryDword(key: string, value: string): number | undefined {
  try {
    const out = execFileSync("reg", ["query", key, "/v", value], {
      encoding: "utf8",
      windowsHide: true,
      timeout: 2_000,
    });
    const match = out.match(new RegExp(value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + "\\s+REG_DWORD\\s+0x([0-9a-f]+)", "i"));
    if (!match) return undefined;
    return Number.parseInt(match[1], 16);
  } catch {
    return undefined;
  }
}

function systemLight(): boolean {
  const apps = registryDword(
    "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize",
    "AppsUseLightTheme",
  );
  return apps !== 0;
}

/** Tab chrome follows the installed Office UI theme, not the Cursor editor palette. */
export function officeTabAppearance(): OfficeTabAppearance {
  const id = registryDword("HKCU\\Software\\Microsoft\\Office\\16.0\\Common", "UI Theme")
    ?? registryDword("HKCU\\Software\\Microsoft\\Office\\15.0\\Common", "UI Theme");
  if (id === 4 || id === 7) {
    return systemLight()
      ? { ...LIGHT, theme: "system-light" }
      : { ...DARK, theme: "system-dark" };
  }
  return BY_ID[id ?? 0] ?? LIGHT;
}
