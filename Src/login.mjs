import puppeteer from "puppeteer";
import sharp from "sharp";
import { execSync } from "node:child_process";
import { createInterface } from "node:readline/promises";
import { stdin, stdout } from "node:process";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import { readdirSync, statSync, existsSync, mkdirSync } from "node:fs";
import { homedir } from "node:os";
import {
  findWindow, isWindow, setForeground, getWindowRect, moveWindow,
  pixelMatch, imageSearch, clickAt, sendCtrlA, sendString, sendEnter,
  moveMouse, getMousePos
} from "./win32.mjs";

const __dirname = dirname(fileURLToPath(import.meta.url));
const EMAIL = process.env.EMAIL;
const TARGET_URL = process.env.TARGET_URL;
const WINDOWS_CREDENTIALS_STORE_VM_CREDENTIALS = process.env.WINDOWS_CREDENTIALS_STORE_VM_CREDENTIALS;
const CITRIX_WINDOW = "W11-STATIC - Desktop Viewer";

if (!EMAIL || !TARGET_URL || !WINDOWS_CREDENTIALS_STORE_VM_CREDENTIALS) {
  console.error("Missing required env vars: EMAIL, TARGET_URL, WINDOWS_CREDENTIALS_STORE_VM_CREDENTIALS. Copy .env.example to .env and fill in values.");
  process.exit(1);
}

function getPassword() {
  const script = join(__dirname, "get-credential.ps1");
  return execSync(
    `powershell -NoProfile -ExecutionPolicy Bypass -File "${script}" "${WINDOWS_CREDENTIALS_STORE_VM_CREDENTIALS}"`,
    { encoding: "utf8" }
  ).trim();
}

async function waitForVisible(page, selector, timeout = 15000) {
  await page.waitForFunction(
    (sel) => {
      const el = document.querySelector(sel);
      return el && el.offsetParent !== null;
    },
    { timeout },
    selector
  );
  await new Promise((r) => setTimeout(r, 500));
}

async function main() {
  console.log("=== Remote Login ===\n");

  // Step 0: Get password from Windows Credential Manager
  console.log("Reading credentials...");
  const password = getPassword();
  console.log(`Password retrieved (${password.length} chars)`);

  const hwnd = findWindow(CITRIX_WINDOW);
  if (hwnd){
    console.log(`Citrix window "${CITRIX_WINDOW}" is already open. Trying to login.`);
    await monitorCitrixLogin(password, null);
  }

  // Pre-load the login-screen template image in parallel with browser launch
  const templatePromise = loadTemplate();

  console.log("Launching browser in incognito mode...");
  const browser = await puppeteer.launch({
    headless: false,
    args: ["--incognito"], // 
    defaultViewport: null,
  });

  const context = await browser.createBrowserContext();
  const page = await context.newPage();
  const pages = await browser.pages();
  for (const p of pages) {
    if (p !== page) await p.close();
  }

  await browser.setWindowBounds(await page.windowId(), {
    left: 5,
    top: 5,
    width: 700,
    height: 1000
  });

  // Set up download directory via CDP so .ica files are saved automatically
  const downloadDir = join(homedir(), "Downloads");
  if (!existsSync(downloadDir)) mkdirSync(downloadDir, { recursive: true });
  const cdp = await page.createCDPSession();
  await cdp.send("Page.setDownloadBehavior", {
    behavior: "allow",
    downloadPath: downloadDir,
  });

  console.log(`Navigating to ${TARGET_URL}...`);
  await page.goto(TARGET_URL, { waitUntil: "networkidle2", timeout: 30000 });

  // Step 1: Fill email
  console.log("Filling email...");
  await page.waitForSelector('input[name="loginfmt"]', { timeout: 15000 });
  await page.type('input[name="loginfmt"]', EMAIL, { delay: 5 });
  await page.click('input[type="submit"]');

  // Step 2: Fill password
  console.log("Filling password...");
  await waitForVisible(page, 'input[name="passwd"]');
  await page.type('input[name="passwd"]', password, { delay: 30 });
  await page.click('input[type="submit"]');

  // Step 3: MFA - choose SMS verification
  // Two variants:
  //   A) Clean run: Authenticator prompt with "I can't use my Microsoft Authenticator app right now" link
  //   B) Retry/error: "Verify your identity" with SMS button directly visible
  console.log("Waiting for MFA page...");
  await page.waitForFunction(
    () => {
      const all = document.body?.innerText || "";
      return all.includes("Approve sign in") ||
             all.includes("Verify your identity") ||
             all.includes("I can't use") ||
             all.includes("Sign in another way");
    },
    { timeout: 20000 }
  );

  await new Promise((r) => setTimeout(r, 1000));
  // Check if SMS button is already visible (variant B), otherwise click the "I can't use" link first (variant A)
  const hasSmsButton = await page.evaluate(() => {
    const buttons = [...document.querySelectorAll("div[role='button'], button")];
    return buttons.some((b) => b.textContent?.includes("Text +"));
  });

  if (!hasSmsButton) {
    console.log('Clicking "I can\'t use my Microsoft Authenticator app right now"...');
    await page.evaluate(() => {
      const links = [...document.querySelectorAll("a, button, [role='button']")];
      const link = links.find((l) => l.textContent?.includes("I can't use") || l.textContent?.includes("Sign in another way"));
      link?.click();
    });

    // Wait for SMS option to appear
    await page.waitForFunction(
      () => {
        const buttons = [...document.querySelectorAll("div[role='button'], button")];
        return buttons.some((b) => b.textContent?.includes("Text +"));
      },
      { timeout: 10000 }
    );
    await new Promise((r) => setTimeout(r, 500));
  }

  console.log("Selecting SMS verification...");
  await page.evaluate(() => {
    const buttons = [...document.querySelectorAll("div[role='button'], button")];
    const smsBtn = buttons.find((b) => b.textContent?.includes("Text +"));
    smsBtn?.click();
  });

  // Step 4: Enter SMS code
  // The user may enter the code directly in the browser, so handle both cases.
  console.log("Waiting for code input...");
  let otcVisible = false;
  try {
    await waitForVisible(page, 'input[name="otc"]', 10000);
    otcVisible = true;
  } catch {
    console.log("Code input not found — assuming code was entered in browser.");
  }

  if (otcVisible) {
    console.log("\nSMS code sent to your phone!");

    const ac = new AbortController();
    const rl = createInterface({ input: stdin, output: stdout });

    const smsGonePromise = page.waitForFunction(
      () => !document.querySelector('input[name="otc"]')?.offsetParent,
      { timeout: 15000 }
    ).catch(err => {
      console.error("SMS wait failed:", err);
    });

    const questionPromise = rl.question("Enter SMS code: ", { signal: ac.signal });

    const raceResult = await Promise.race([
      questionPromise
        .then((ans) => ({ type: "code", value: ans.trim() }))
        .catch((err) => (err.name === "AbortError" ? { type: "aborted" } : Promise.reject(err))),
      smsGonePromise.then(() => ({ type: "gone" })),
    ]);
    ac.abort();
    rl.close();

    if (raceResult.type === "code" && raceResult.value) {
      const stillVisible = await page.evaluate(() => {
        const el = document.querySelector('input[name="otc"]');
        return el && el.offsetParent !== null;
      });
      if (stillVisible) {
        await page.type('input[name="otc"]', raceResult.value, { delay: 60 });
        await page.click('input[type="submit"]');
      }
    } else {
      console.log("Code input disappeared — continuing (code was entered in browser).");
    }
  }

  // Step 5: Handle "Stay signed in?" prompt if it appears
  try {
    await page.waitForSelector("#KmsiDescription", { timeout: 4000 });
    console.log('Handling "Stay signed in?" prompt...');
    await page.click('input[type="submit"]');
  } catch {
    // No "Stay signed in" prompt — continue
  }

  // Step 6: Citrix Workspace - wait for it to load
  console.log("Waiting for Citrix Workspace...");
  await page.waitForFunction(
    () => document.title?.includes("Citrix"),
    { timeout: 60000 }
  );
  // Wait until loading indicator is gone and DESKTOPS tab is clickable
  console.log("Waiting for Citrix Workspace to finish loading...");
  await page.waitForFunction(
    () => {
      // Check no loading spinner/overlay is visible
      const loader = document.querySelector(".loader, .loading, [class*='loading'], [class*='spinner'], #mask");
      if (loader && loader.offsetParent !== null) return false;
      // Check DESKTOPS tab is present
      const tabs = [...document.querySelectorAll("[role='tab'], .tab, a")];
      return tabs.some((t) => t.textContent?.includes("DESKTOPS") || t.textContent?.includes("Desktops"));
    },
    { timeout: 60000 }
  );
  await new Promise((r) => setTimeout(r, 500));

  // Step 7: Click DESKTOPS tab
  console.log("Clicking DESKTOPS tab...");
  await page.evaluate(() => {
    const tabs = [...document.querySelectorAll("[role='tab'], .tab, a")];
    const dt = tabs.find((t) => t.textContent?.includes("DESKTOPS") || t.textContent?.includes("Desktops"));
    dt?.click();
  });
  await new Promise((r) => setTimeout(r, 1000));

  // Step 8: Click "W11-STATIC" desktop
  console.log('Clicking "W11-STATIC" desktop...');
  await page.waitForFunction(
    () => {
      const links = [...document.querySelectorAll("a, [role='link'], .storeapp-icon, img[alt]")];
      return links.some((l) =>
        l.textContent?.includes("W11-STATIC") ||
        l.alt?.includes("W11-STATIC")
      );
    },
    { timeout: 10000 }
  );
  // Snapshot existing .ica files (name → mtimeMs) before clicking
  const icaBefore = new Map();
  for (const f of readdirSync(downloadDir).filter((n) => n.endsWith(".ica"))) {
    icaBefore.set(f, statSync(join(downloadDir, f)).mtimeMs);
  }

  await page.evaluate(() => {
    const links = [...document.querySelectorAll("a, [role='link'], img[alt]")];
    const desktop = links.find((l) =>
      l.textContent?.includes("W11-STATIC") ||
      l.alt?.includes("W11-STATIC")
    );
    desktop?.click();
  });

  // Step 9: Wait for .ica file (new file or same name with updated mtime)
  console.log("Waiting for .ica file download...");
  const icaFile = await new Promise((resolve, reject) => {
    const start = Date.now();
    const check = () => {
      for (const f of readdirSync(downloadDir).filter((n) => n.endsWith(".ica"))) {
        const full = join(downloadDir, f);
        const mtime = statSync(full).mtimeMs;
        if (!icaBefore.has(f) || mtime > icaBefore.get(f)) {
          return resolve(full);
        }
      }
      if (Date.now() - start > 15000) return reject(new Error("ICA download timeout"));
      setTimeout(check, 500);
    };
    check();
  });

  console.log(`Downloaded: ${icaFile}`);
  console.log("Launching Citrix session...");
  execSync(`start "" "${icaFile}"`, { shell: "cmd.exe" });

  console.log("\n✓ Login complete! Citrix desktop session launching.");
  console.log("  Browser will close automatically in 5 seconds.\n");
  //await new Promise((r) => setTimeout(r, 5000));
  //await browser.close();

  // Step 10: Monitor Citrix VM and auto-login (template already pre-loaded)
  const template = await templatePromise;
  await monitorCitrixLogin(password, template);

  console.log("Done!");
}

/** Load the reference login-screen icon as raw RGBA pixel data. */
async function loadTemplate() {
  const imgPath = join(__dirname, "images-to-find", "login-screen-icon.png");
  if (!existsSync(imgPath)) {
    console.log(`\n⚠  Reference image not found: ${imgPath}`);
    console.log("   Citrix VM auto-login will be skipped.");
    return null;
  }
  const img = sharp(imgPath);
  const meta = await img.metadata();
  const buf = await img.ensureAlpha().raw().toBuffer();
  return { width: meta.width, height: meta.height, data: buf };
}

/**
 * Replicates the AHK monitoring loop:
 * - Waits for the Citrix Desktop Viewer window
 * - Checks a pixel to confirm the window is loaded
 * - Resizes +1px to trigger a redraw
 * - Searches for the login screen icon and types the password
 * - Resizes back and waits for the window to close
 */
async function monitorCitrixLogin(password, template) {
  if (!template) {
    template = await loadTemplate();
    if (!template) return;
  }

  console.log(`\n=== Citrix VM Login Monitor ===`);
  console.log(`Watching for "${CITRIX_WINDOW}"...\n`);

  // AHK pixel check: coordinate (12, 37), color 0x3A4145 (RGB) → R=0x3A, G=0x41, B=0x45
  const PIXEL_X = 12, PIXEL_Y = 37;
  const PIX_R = 0x3a, PIX_G = 0x41, PIX_B = 0x45;
  // AHK ImageSearch region (window coords)
  const SEARCH_X = 917, SEARCH_Y = 400, SEARCH_X2 = 1459, SEARCH_Y2 = 900;

  let attempt = 0;
  for (;;) {
    attempt++;
    await sleep(500);

    const hwnd = findWindow(CITRIX_WINDOW);
    if (!hwnd){
      console.log(`Citrix window is not there yet... Waiting ${attempt}.`);
      continue;
    }

    // Bring window to foreground and wait until it's actually active
    const fgOk = setForeground(hwnd);
    if (!fgOk) {
      console.log("Could not bring Citrix window to foreground — retrying...");
      continue;
    }

    // Window exists — wait until it's ready (pixel check), but skip after 3s
    // let pixelOk = false;
    // for (let i = 0; i < 10; i++) {
    //   if (pixelMatch(hwnd, PIXEL_X, PIXEL_Y, PIX_R, PIX_G, PIX_B, 15)) {
    //     pixelOk = true;
    //     break;
    //   }
    //   console.log(`Pixel check does not match yet... Waiting (${i + 1}/10).`);
    //   await sleep(500);
    // }
    // if (!pixelOk) {
    //   console.log("Pixel check timed out — continuing anyway.");
    // }
    console.log("Citrix window detected and loaded.");

    // Resize +1px width (triggers a redraw, same as AHK)
    const rect = getWindowRect(hwnd);
    moveWindow(hwnd, rect.x, rect.y, rect.w + 1, rect.h);

    // Save mouse position and restore later
    const origMouse = getMousePos();
    console.log("Saved original mouse position:", origMouse);

    // Click inside the VM to activate it (needed for image to render)
    const centerX = rect.x + Math.floor(rect.w / 2);
    const centerY = rect.y + Math.floor(rect.h / 2);
    clickAt(centerX, centerY);
    await sleep(150);

    // Search for login screen icon
    console.log("Searching for login screen...");
    let found = null;
    for (let attempt = 0; attempt < 60; attempt++) {
      const sw = SEARCH_X2 - SEARCH_X;
      const sh = SEARCH_Y2 - SEARCH_Y;
      found = imageSearch(hwnd, template, SEARCH_X, SEARCH_Y, sw, sh, 15);
      if (found) break;
      await sleep(200);
      // Abort if window disappeared
      if (!isWindow(hwnd)) break;
    }

    if (!found) {
      console.log("Login screen not found within timeout — will keep watching.");
      // Resize back
      moveWindow(hwnd, rect.x, rect.y, rect.w, rect.h);
      continue;
    }

    console.log(`Login icon found at (${found.x}, ${found.y}). Typing credentials...`);

    // Convert window coords to screen coords for clicking
    const curRect = getWindowRect(hwnd);
    const screenX = curRect.x + found.x;
    const screenY = curRect.y + found.y;
    console.log(`  Screen click target: (${screenX}, ${screenY})`);

    // Click on the login icon (password field) — click twice with delay for Citrix RDP latency
    clickAt(screenX, screenY);
    await sleep(200);
    clickAt(screenX, screenY);
    await sleep(200);

    // Restore mouse position to avoid interfering with user
    if (origMouse) {
      console.log(`Restoring mouse position`);
      moveMouse(origMouse.x, origMouse.y);
    }

    // Type password via scan codes + Enter
    sendString(password);
    await sleep(200);
    sendEnter();

    // Resize back to original
    await sleep(1000);
    moveWindow(hwnd, rect.x, rect.y, rect.w, rect.h);

    // Wait for the Citrix window to close
    console.log("Waiting for Citrix session to close...");
    while (isWindow(hwnd)) {
      await sleep(5000);
    }
    console.log("Citrix window closed.");
    // Loop back to watch for a new session
  }
}

function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

main().catch((err) => {
  console.error("Error:", err.message);
  process.exit(1);
});
