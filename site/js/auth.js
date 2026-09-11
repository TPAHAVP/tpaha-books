// Sign-in wrapper around the vendored MSAL UMD build (window.msal). An empty clientId means local mode.
const SCOPES = ['Files.ReadWrite', 'User.Read'];

export async function initAuth(config) {
  if (!config.clientId) {
    return {
      mode: 'local', account: null,
      signIn: async () => {}, signOut: async () => {},
      getToken: async () => { throw new Error('Local mode has no OneDrive access.'); },
    };
  }
  const msal = globalThis.msal;
  if (!msal) throw new Error('The Microsoft sign-in library failed to load.');
  const redirectUri = new URL('./', location.href).href;   // site root; must match the Entra app registration
  const pca = new msal.PublicClientApplication({
    auth: {
      clientId: config.clientId,
      authority: `https://login.microsoftonline.com/${config.tenantId || 'organizations'}`,
      redirectUri,
      postLogoutRedirectUri: redirectUri,
    },
    cache: { cacheLocation: 'localStorage' },
  });
  await pca.initialize();
  let account = null;
  try {
    const result = await pca.handleRedirectPromise({ hash: location.hash, navigateToLoginRequestUrl: true });
    if (result && result.account) account = result.account;
  } catch (e) {
    console.error('Sign-in redirect failed', e);
  }
  account = account || pca.getActiveAccount() || pca.getAllAccounts()[0] || null;
  if (account) pca.setActiveAccount(account);

  const interactionNeeded = e => (msal.InteractionRequiredAuthError && e instanceof msal.InteractionRequiredAuthError)
    || /interaction_required|login_required|consent_required|no_tokens_found|invalid_grant|monitor_window_timeout/i.test((e && (e.errorCode || e.message)) || '');

  return {
    mode: 'onedrive',
    account: account ? { name: account.name || account.username, username: account.username, homeAccountId: account.homeAccountId, tenantId: account.tenantId, localAccountId: account.localAccountId } : null,
    tenantId: config.tenantId || (account && account.tenantId) || '',
    async signIn() {
      await pca.loginRedirect({ scopes: SCOPES, redirectStartPage: location.href, prompt: 'select_account' });
    },
    async signOut() {
      await pca.logoutRedirect({ account, postLogoutRedirectUri: redirectUri });
    },
    async getToken() {
      if (!account) throw Object.assign(new Error('Not signed in'), { code: 'not_signed_in', status: 401 });
      try {
        return (await pca.acquireTokenSilent({ scopes: SCOPES, account })).accessToken;
      } catch (e) {
        if (interactionNeeded(e)) {
          await pca.acquireTokenRedirect({ scopes: SCOPES, account, redirectStartPage: location.href });
          return new Promise(() => {});   // page navigates away
        }
        throw e;
      }
    },
  };
}
