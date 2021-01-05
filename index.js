import "https://alcdn.msauth.net/lib/1.4.4/js/msal.js";
import {
  html,
  render,
} from "https://unpkg.com/htm/preact/standalone.module.js";

function createMsal(tenantId) {
  return new Msal.UserAgentApplication({
    auth: {
      authority: `https://login.microsoftonline.com/${tenantId}`,
      clientId: "203f1145-856a-4232-83d4-a43568fba23d",
      redirectUri: "https://dataexplorer-dev.azurewebsites.net",
    },
  });
}

let msal = createMsal("common");

async function switchTenant(tenantId) {
  console.log(`Switching to Tenant: ${tenantId}`);
  msal = createMsal(tenantId);
  const response = await msal.loginPopup();
  const tokenResponse = await msal.acquireTokenSilent({
    scopes: ["https://management.azure.com//.default"],
  });
  const token = tokenResponse.accessToken;

  const tenants = await fetchTenants(token);
  const subs = await fetchSubscriptions(token);

  reRender(tenants, subs, response.tenantId);
}

async function login() {
  const response = await msal.loginPopup();
  const tokenResponse = await msal.acquireTokenSilent({
    scopes: ["https://management.azure.com//.default"],
  });
  const token = tokenResponse.accessToken;
  const tenants = await fetchTenants(token);
  const subs = await fetchSubscriptions(token);

  reRender(tenants, subs, response.tenantId);
}

function reRender(tenants, subs, tenantId) {
  render(
    html`<p>Tenant - ${tenantId}</p>
      <ul>
        ${tenants.map(
          (t) =>
            html`<li>
              ${t.displayName} - ${t.tenantId}
              <button onClick=${() => switchTenant(t.tenantId)}>Switch</button>
            </li>`
        )}
      </ul>
      <p>Subscriptions</p>
      <ul>
        ${subs.map((s) => html`<li>${s.displayName} - ${s.id}</li>`)}
      </ul>`,
    document.getElementById("data")
  );
}

render(html`<button onClick=${login}>Sign In</button>`, document.body);

async function fetchTenants(accessToken) {
  const headers = new Headers();
  const bearer = `Bearer ${accessToken}`;

  headers.append("Authorization", bearer);

  let tenents = [];
  let nextLink = `https://management.azure.com/tenants?api-version=2020-01-01`;

  while (nextLink) {
    const response = await fetch(nextLink, { headers });
    const result =
      response.status === 204 || response.status === 304
        ? undefined
        : await response.json();
    if (!response.ok) {
      throw result;
    }
    nextLink = result.nextLink;
    tenents = [...tenents, ...result.value];
  }
  return tenents;
}

export async function fetchSubscriptions(accessToken) {
  const headers = new Headers();
  const bearer = `Bearer ${accessToken}`;

  headers.append("Authorization", bearer);

  let subscriptions = [];
  let nextLink = `https://management.azure.com/subscriptions?api-version=2020-01-01`;

  while (nextLink) {
    const response = await fetch(nextLink, { headers });
    const result =
      response.status === 204 || response.status === 304
        ? undefined
        : await response.json();
    if (!response.ok) {
      throw result;
    }
    nextLink = result.nextLink;
    const validSubscriptions = result.value.filter(
      (sub) =>
        sub.state === "Enabled" ||
        sub.state === "Warned" ||
        sub.state === "PastDue"
    );
    subscriptions = [...subscriptions, ...validSubscriptions];
  }
  return subscriptions;
}
