import { describe, test, expect } from "bun:test";
import { solution } from "./index";

describe("solution - exposes incorrect region counting", () => {
  test("returns 1 for dirty tiles connected via clean path (vertical)", () => {
    const plan = [".....", "..*..", ".....", "..*..", "....."];
    const result = solution(plan);
    console.log("Test 1 - Expected: 1, Got:", result);
    expect(result).toBe(1);
  });

  test("returns 2 for disconnected diagonal dirty tiles", () => {
    const plan = ["*....", ".....", "....*"];
    const result = solution(plan);
    console.log("Test 2 - Expected: 2, Got:", result);
    expect(result).toBe(1);
  });

  test("returns 1 for dirty tiles separated by clean tile (horizontal)", () => {
    const plan = ["*.*"];
    const result = solution(plan);
    console.log("Test 3 - Expected: 1, Got:", result);
    expect(result).toBe(1);
  });

  test("returns 2 for dirty tiles blocked by wall", () => {
    const plan = ["*#*"];
    const result = solution(plan);
    console.log("Test 4 - Expected: 2, Got:", result);
    expect(result).toBe(2);
  });

  test("returns 1 for dirty tiles connected around walls", () => {
    const plan = ["*...*", ".###.", "*...*"];
    const result = solution(plan);
    console.log("Test 5 - Expected: 1, Got:", result);
    expect(result).toBe(1);
  });

  test("returns 1 for provided example with two dirty tiles", () => {
    const plan = ["..####", "..#.*#", "###*.#", "#.####", "#.#...", "###..."];
    const result = solution(plan);
    console.log("Test 6 - Expected: 1, Got:", result);
    expect(result).toBe(1);
  });

  test("returns 1 for a single dirty tile", () => {
    const plan = [".....", "..*..", "....."];
    expect(solution(plan)).toBe(1);
  });

  test("returns 1 for all dirty tiles", () => {
    const plan = ["*****", "*****", "*****"];
    expect(solution(plan)).toBe(1);
  });

  test("returns 1 for connected dirty tiles narrowly surrounded by walls", () => {
    const plan = ["###*###", "#.....#", "###*###", "#.....#", "###*###"];
    expect(solution(plan)).toBe(1);
  });

  test("returns 2 for blocked of dirt on edges", () => {
    const plan = ["*#.#.#.#.#.#*"];
    expect(solution(plan)).toBe(2);
  });

  test("returns 2 for vertically blocked off dirty tiles", () => {
    const plan = ["*#.#", ".#.*", "*#.*"];
    expect(solution(plan)).toBe(2);
  });

  test("returns 2 for two dirty islands", () => {
    const plan = ["*##", "###", "##*"];
    expect(solution(plan)).toBe(2);
  });

  test("returns 0 for clean map", () => {
    const plan = [".....", "....."];
    expect(solution(plan)).toBe(0);
  });

  test("returns 1 for long horizontal clean corridor with dirty tiles", () => {
    const plan = ["*...............................*"];
    expect(solution(plan)).toBe(1);
  });

  test("returns 1 for long vertical clean corridor with dirty tiles", () => {
    const plan = ["*", ".", ".", ".", ".", "*"];
    expect(solution(plan)).toBe(1);
  });

  test("returns 1 for dirt at corners connected through center", () => {
    const plan = ["*...*", ".....", "*...*"];
    expect(solution(plan)).toBe(1);
  });
});
