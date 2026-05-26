// ใส่ URL Web App ของ Google Apps Script ตรงนี้
// ต้องเป็นลิงก์ที่ลงท้ายด้วย /exec
const GAS_URL = "https://script.google.com/macros/s/AKfycbxKRr0x6oF2vuzinXAeuEqQXck1PMUYFQWVbrm-Q8nQlj_sGSOnLEwdrvZ2UIb4j3ju/exec";

const setupBox = document.getElementById("setupBox");
const appBox = document.getElementById("appBox");
const gasFrame = document.getElementById("gasFrame");
const openDirect = document.getElementById("openDirect");

const isConfigured =
  GAS_URL &&
  GAS_URL.startsWith("https://script.google.com/macros/s/") &&
  GAS_URL.endsWith("/exec");

if (!isConfigured) {
  setupBox.classList.remove("hidden");
  openDirect.style.display = "none";
} else {
  gasFrame.src = GAS_URL;
  openDirect.href = GAS_URL;
  appBox.classList.remove("hidden");
}
