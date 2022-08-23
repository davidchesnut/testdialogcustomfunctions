/* global clearInterval, console, CustomFunctions, setInterval, OfficeRuntime */

import { resolveProjectReferencePath } from "typescript";

/**
 * Displays the current time once a second.
 * @customfunction
 * @param invocation Custom function handler
 */
export function clock(invocation: CustomFunctions.StreamingInvocation<string>): void {
  const timer = setInterval(() => {
    const time = currentTime();
    invocation.setResult(time);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Returns the current time.
 * @returns String with the current time formatted for the current locale.
 */
export function currentTime(): string {
  return new Date().toLocaleTimeString();
}

/**
 * Increments a value once a second.
 * @customfunction
 * @param incrementBy Amount to increment
 * @param invocation Custom function handler
 */
export function increment(incrementBy: number, invocation: CustomFunctions.StreamingInvocation<number>): void {
  let result = 0;
  const timer = setInterval(() => {
    result += incrementBy;
    invocation.setResult(result);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Writes a message to console.log().
 * @customfunction LOG
 * @param message String to write.
 * @returns String to write.
 */
export function logMessage(message: string): Promise<string> {
  console.log(message);
  return new Promise(function (resolve, reject) {
    OfficeRuntime.displayWebDialog("https://localhost:3000/dialog.html", {
      height: "50%",
      width: "50%",
      onMessage: function (answer, dialog) {
        console.log(answer);
        dialog.close();
        resolve(answer);
      },
      onRuntimeError: function (error, dialog) {
        dialog.close();
        reject(error);
      },
    }).catch(function (e) {
      reject(e);
    });
  });
}
