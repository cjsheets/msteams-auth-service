export function insideIframe() {
  try {
    return window.self !== window.top;
  } catch (e) {
    return true;
  }
}
