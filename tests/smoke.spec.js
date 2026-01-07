const { test, expect } = require("@playwright/test");

test("app loads", async ({ page }) => {
  await page.goto("/");
  await expect(page.locator(".title")).toHaveText("[MULTI]THREADER");
  await expect(page.locator("#threadTabs")).toBeVisible();
  await expect(page.locator("#sheetTable")).toBeVisible();
});
