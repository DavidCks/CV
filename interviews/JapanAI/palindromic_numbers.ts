/**
 *
 * @param S = String of decimal digits (0-9)
 *
 * @description
 * Creates a palindromic number with the largest possible decimal value,
 * using at least one digit from S, by reordering the digits from S into it's palindromic numbers
 * and picking the largest one.
 *
 * A palindromic number is a number that remains the same when it's digits are reversed:
 * 323 -> 323 (ok)
 * 123 -> 321 (nope)
 *
 * leading and trailing 0s are stripped
 *
 * @return String: A palindromic number with the largest possible decimal value
 */
export function solution(S: string): string {
  const freq: number[] = new Array(10).fill(0);

  for (const char of S) {
    const digit = parseInt(char);
    if (digit >= 0 && digit <= 9) {
      freq[digit]!++;
    }
  }

  let left = "";
  let middle = "";

  for (let digit = 9; digit >= 0; digit--) {
    const pairs = Math.floor(freq[digit]! / 2);
    left += String(digit).repeat(pairs);
    freq[digit]! %= 2;
  }

  for (let digit = 9; digit >= 0; digit--) {
    if (freq[digit]! > 0) {
      middle = String(digit);
      break;
    }
  }

  const right = [...left].reverse().join("");
  let result = left + middle + right;

  result = result.replace(/^0+|0+$/g, "");
  return result === "" ? "0" : result;
}
