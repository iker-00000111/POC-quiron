import {
  PublicClientApplication,
  InteractionRequiredAuthError,
} from "https://cdn.jsdelivr.net/npm/@azure/msal-browser@4.11.0/+esm";

const APP_CONFIG = {
  clientId: "27731670-72d7-455a-a1c1-c52e75c0cfd3",
  tenantId: "6f167842-765d-4bda-a56e-9d433bc27739",
  apiScope: "api://dcde18b5-ab34-437e-b3b9-fa080940e4f8/access_as_user",
  apiBaseUrl: "https://api-quiron-back.azurewebsites.net",
};

const authority = `https://login.microsoftonline.com/${APP_CONFIG.tenantId}`;
const msalInstance = new PublicClientApplication({
  auth: {
    clientId: APP_CONFIG.clientId,
    authority,
    redirectUri: window.location.origin,
  },
  cache: {
    cacheLocation: "sessionStorage",
  },
});

let isAuthenticated = false;
let loading = false;
let citas = [];
let currentAccount = null;

const loginScreen = document.getElementById("loginScreen");
const appScreen = document.getElementById("appScreen");
const loadingText = document.getElementById("loadingText");
const statusText = document.getElementById("statusText");
const tableWrapper = document.getElementById("tableWrapper");
const citasBody = document.getElementById("citasBody");
const loginButton = document.getElementById("loginButton");
const logoutButton = document.getElementById("logoutButton");
const welcomeText = document.querySelector(".welcome");

function renderCitas() {
  console.log("👉 Ejecutando renderCitas. ¿Cuántas citas hay corporativas?:", citas.length);
  
  if (!citas.length) {
    citasBody.innerHTML = `
      <tr>
        <td colspan="3">No hay datos para mostrar.</td>
      </tr>
    `;
    return;
  }

  // Mapeamos los campos exactos que te ha escupido la consola: Id, Nombre, Email
  citasBody.innerHTML = citas
    .map(
      (cita) => `
        <tr>
          <td>${cita.Id}</td>
          <td>${cita.Nombre}</td>
          <td>${cita.Email}</td>
        </tr>
      `
    )
    .join("");
    
  console.log("✅ HTML inyectado con éxito en el tbody.");
}

function setStatus(message = "", type = "info") {
  if (!message) {
    statusText.hidden = true;
    statusText.textContent = "";
    statusText.className = "status info";
    return;
  }
  statusText.hidden = false;
  statusText.textContent = message;
  statusText.className = `status ${type}`;
}

function render() {
  loginScreen.hidden = isAuthenticated;
  appScreen.hidden = !isAuthenticated;
  loadingText.hidden = !loading;
  tableWrapper.hidden = loading;
  welcomeText.textContent = currentAccount
    ? `Bienvenido, ${currentAccount.name || "Usuario"}`
    : "Bienvenido, Usuario";
}

async function getApiToken() {
  const request = {
    scopes: [APP_CONFIG.apiScope],
    account: currentAccount,
  };

  try {
    const response = await msalInstance.acquireTokenSilent(request);
    return response.accessToken;
  } catch (error) {
    if (error instanceof InteractionRequiredAuthError) {
      const popupResponse = await msalInstance.acquireTokenPopup(request);
      return popupResponse.accessToken;
    }
    throw error;
  }
}

async function fetchCitas() {
  const token = await getApiToken();
  const response = await fetch(`${APP_CONFIG.apiBaseUrl}/api/datos`, {
    method: "GET",
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json",
    },
  });

  if (!response.ok) {
    const errorBody = await response.text();
    throw new Error(
      `Error ${response.status} al consultar la API: ${errorBody || "sin detalle"}`
    );
  }

  return response.json();
}

async function login() {
  setStatus();
  isAuthenticated = true;
  loading = true;
  citas = [];
  render();

  try {
    const loginResponse = await msalInstance.loginPopup({
      scopes: ["openid", "profile", "email"],
    });

    currentAccount = loginResponse.account;
    const apiData = await fetchCitas();
    
    // 🌟 AÑADE ESTA LÍNEA AQUÍ PARA VER LOS DATOS EN LA CONSOLA (F12)
    console.log("¡DATOS REALES DE CLIENTES!", apiData);

    citas = Array.isArray(apiData) ? apiData : [];
    renderCitas();
    setStatus("Datos cargados correctamente.", "info");
  } catch (error) {
    isAuthenticated = false;
    currentAccount = null;
    citas = [];
    setStatus(error.message || "No se pudo iniciar sesión.", "error");
  } finally {
    loading = false;
    render();
  }
}

async function logout() {
  setStatus();
  try {
    if (currentAccount) {
      await msalInstance.logoutPopup({
        account: currentAccount,
        mainWindowRedirectUri: window.location.origin,
      });
    }
  } catch (error) {
    setStatus(error.message || "Error al cerrar sesión.", "error");
  }

  isAuthenticated = false;
  loading = false;
  citas = [];
  currentAccount = null;
  render();
}

loginButton.addEventListener("click", login);
logoutButton.addEventListener("click", logout);

await msalInstance.initialize();
await msalInstance.handleRedirectPromise();
const accounts = msalInstance.getAllAccounts();
if (accounts.length > 0) {
  currentAccount = accounts[0];
  isAuthenticated = true;
  loading = true;
  render();
  try {
    const apiData = await fetchCitas();
    citas = Array.isArray(apiData) ? apiData : [];
    renderCitas();
    setStatus("Sesión recuperada y datos actualizados.", "info");
  } catch (error) {
    setStatus(error.message || "No se pudieron recuperar los datos.", "error");
  } finally {
    loading = false;
    render();
  }
} else {
  render();
}
