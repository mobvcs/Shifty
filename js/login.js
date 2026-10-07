import { getSession, saveSession, verifyCredentials } from "./auth.js";

const APP_URL = "./app.html";

// 이미 로그인된 상태면 바로 앱으로 이동
if (getSession()) {
  window.location.replace(APP_URL);
}

const form = document.getElementById("loginForm");
const userInput = document.getElementById("loginUsername");
const pwInput = document.getElementById("loginPassword");
const errorEl = document.getElementById("loginError");
const submitBtn = document.getElementById("loginSubmitBtn");

userInput.focus();

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  errorEl.textContent = "";
  submitBtn.disabled = true;
  try {
    const account = await verifyCredentials(userInput.value, pwInput.value);
    if (!account) {
      errorEl.textContent = "아이디 또는 비밀번호가 올바르지 않습니다.";
      pwInput.value = "";
      pwInput.focus();
      return;
    }
    saveSession(account.username, account.level);
    window.location.replace(APP_URL);
  } catch (error) {
    errorEl.textContent = error.message;
  } finally {
    submitBtn.disabled = false;
  }
});
