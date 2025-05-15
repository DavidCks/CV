function isValid(
  plan: string[],
  visited: boolean[][],
  x: number,
  y: number
): boolean {
  return (
    x >= 0 &&
    x < plan.length &&
    y >= 0 &&
    y < plan[x]!.length &&
    plan[x]![y] !== "#" &&
    !visited[x]![y]
  );
}

function bfs(plan: string[], visited: boolean[][], x: number, y: number): void {
  const dx: number[] = [0, 0, 1, -1];
  const dy: number[] = [1, -1, 0, 0];

  const q: [number, number][] = [];
  q.push([x, y]);
  while (q.length > 0) {
    const v = q.shift()!;
    for (let i = 0; i < 4; i++) {
      const nx = v[0] + dx[i]!,
        ny = v[1] + dy[i]!;
      if (isValid(plan, visited, nx, ny)) {
        visited[nx]![ny] = true;
        q.push([nx, ny]);
      }
    }
  }
}

export function solution(plan: string[]): number {
  const n = plan.length,
    m = plan[0]!.length;
  const visited: boolean[][] = new Array<boolean[]>(n);
  for (let i = 0; i < n; i++) {
    visited[i] = new Array<boolean>(m);
  }

  let robots = 0;
  for (let i = 0; i < n; i++) {
    for (let j = 0; j < m; j++) {
      if (plan[i]![j] === "*" && !visited[i]![j]) {
        bfs(plan, visited, i, j);
        robots++;
      }
    }
  }
  return robots;
}
