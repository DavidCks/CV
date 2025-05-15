import { describe, test, expect } from "bun:test";
import { solution } from "./palindromic_numbers";

describe("solution - largest palindromic number", () => {
  test("returns 454 for '543345'", () => {
    expect(solution("543345")).toBe("543345");
  });

  test("returns 3 for '123'", () => {
    expect(solution("123")).toBe("3");
  });

  test("returns 1 for '1'", () => {
    expect(solution("1")).toBe("1");
  });

  test("returns 11 for '11'", () => {
    expect(solution("11")).toBe("11");
  });

  test("returns 121 for '112'", () => {
    expect(solution("112")).toBe("121");
  });

  test("returns 0 for empty input", () => {
    expect(solution("")).toBe("0");
  });

  test("returns 0 for all zeros", () => {
    expect(solution("0000")).toBe("0");
  });

  test("returns 1 for '0001'", () => {
    expect(solution("0001")).toBe("1");
  });

  test("returns 1001 for '1010'", () => {
    expect(solution("1010")).toBe("1001");
  });

  test("returns 543212345 for double digits up to 5", () => {
    expect(solution("122334455")).toBe("543212345");
  });

  test("returns 21212 for '111222'", () => {
    expect(solution("111222")).toBe("21212");
  });

  test("returns 1111 for '1111'", () => {
    expect(solution("1111")).toBe("1111");
  });

  test("returns 11211 for '11211'", () => {
    expect(solution("11211")).toBe("11211");
  });

  test("returns 999 for '9992'", () => {
    expect(solution("9992")).toBe("999");
  });

  // permutations of 99922
  test("returns 99922 for '99922'", () => {
    expect(solution("99922")).toBe("92929");
  });

  test("returns 92929 for '22999'", () => {
    expect(solution("22999")).toBe("92929");
  });

  test("returns 92929 for '29299'", () => {
    expect(solution("29299")).toBe("92929");
  });

  test("returns 92929 for '92299'", () => {
    expect(solution("92299")).toBe("92929");
  });

  test("returns 9 for '9876543210'", () => {
    expect(solution("9876543210")).toBe("9");
  });

  test("returns 9 for '123456789'", () => {
    expect(solution("123456789")).toBe("9");
  });

  test("returns 1 for '1'", () => {
    expect(solution("1")).toBe("1");
  });
});
