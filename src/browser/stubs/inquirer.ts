// replaces the @inquirer/* packages. The browser host replaces the functions
// that use them, so this is never called.
export default function prompt(): never {
  throw new Error('Prompts are not available in the browser build');
}
