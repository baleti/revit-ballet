# Agent engine benchmarks

Latency, throughput, memory and speed of the three script engines (`roslyn` C#, `ironpython2`, `cpython3`) running inside a live Revit.
Harness: `tests/bench/` (`bench.ps1`, `cold.ps1`, one task per language under `tasks/`). Measured 2026-10-10.

## Setup

- Revit 2026.1 (26.1.0.35) in the offline test VM: 10 vCPU, 32 GB RAM, GTX 1060 passed through. RDP session disconnected during the run (no display).
- Model: *Snowdon Towers Sample Architectural* (37,867 elements, opened through the agent API in about 30 s).
- Engines: Roslyn 4.8.0, IronPython 2.7.12, CPython 3.12.8 (pythonnet 3.0).
- The client runs inside the guest (PowerShell + HttpClient), so the numbers include the real HTTPS round trip and no network.
- Per task and engine: 2 warm-up runs, then 7 measured runs; medians reported. The same logic is implemented in all three languages and every engine
  must return the same value (checked: all agree).
- "wall" = client-side time for one job. The server splits it into **compile** (worker thread), **dispatch** (waiting for the UI thread) and **run** (UI thread).
  The HTTP round trip itself is about 9 ms; dispatch to the UI thread is about 3 ms for every engine.

## Latency: one job, warm (milliseconds, median wall = compile + run)

| task | what it does | roslyn | ironpython2 | cpython3 |
|---|---|---:|---:|---:|
| noop | `return 1` (the floor) | 299 (285+3) | **12** (0+2) | **10** (0+0) |
| loop | 3,000,000 iterations of integer arithmetic | 335 (295+30) | 1922 | 568 |
| fib | recursive fib(27), about 830k calls | 316 (290+7) | 518 | **56** |
| collections | build, sort and join 300k ints | 386 (321+55) | 1073 | **112** |
| xyz | 50,000 `XYZ` objects + `DistanceTo` (API interop, no model) | 352 (306+9) | **64** | 523 |
| api-read | all 37,867 elements: id, category name | 482 (345+135) | **261** | 340 |
| api-params | all parameters of 3,000 elements | 999 (326+659) | **768** | 1214 |
| api-write | set a parameter on 2,000 elements in a transaction, roll back | 483 (355+116) | **201** | 322 |
| output | 20,000 lines of output | 373 (298+55) | **57** | 116 |
| value | return a list of 20,000 ints | 337 (309+3) | **14** | 38 |

Roslyn's run time is by far the best (fib 7 ms, loop 30 ms), but every Roslyn job pays a compile of about 300 ms before it runs anything.

## Throughput (100 trivial jobs, jobs per second)

| | roslyn | ironpython2 | cpython3 |
|---|---:|---:|---:|
| one after another | 2.7 | 80 | 92 |
| submitted all at once | 2.7 | 74 | 75 |

## Cold start (fresh Revit, first job of the engine)

| | first job wall | second job wall | Revit working set cost |
|---|---:|---:|---:|
| roslyn | 1.8 s (compile 1.06 s) | 447 ms | **+880 MB** (697 to 1578 MB) |
| ironpython2 | 690 ms | 12 ms | +12 MB |
| cpython3 | 1.4 s (interpreter start-up) | 9 ms | +50 MB |

## What the numbers say

- **Short, interactive jobs** (agents asking a question, setting a value): the Python engines answer in 10 to 15 ms, Roslyn in about 300 ms.
- **Pure computation**: CPython 3 is the fastest Python (fib 10x faster than IronPython); Roslyn is 10 to 70x faster than either, once compiled.
- **Revit API loops**: the cost is dominated by the API itself. IronPython is consistently the best of the three there because it calls .NET directly;
  pythonnet (CPython) pays a marshalling cost per call, visible in `xyz` (10x IronPython, 55x C#).
- **Choosing**: C# for heavy computation or when a typed API helps; IronPython 2 for API-heavy scripts and low-latency agent chatter; CPython 3 for
  pure-Python logic and the modern Python ecosystem, accepting slower API calls.

## Problems the benchmark found (and what was done)

1. **Memory grew without limit under Roslyn (fixed).** After about 350 jobs Revit's working set had reached 25 GB on a 32 GB machine. Cause: every finished
   job kept its compiled script (and with it the whole Roslyn compilation) alive for as long as it stayed in the list of the last 100 jobs, and each
   compilation holds a few hundred MB of native memory the garbage collector cannot see. Fix: a finished job drops its compiled script
   (`AgentHost.Finish`), the Roslyn engine shares one assembly loader (it previously created one load context per script) and collects garbage every
   fourth compile. The working set now stays between 1.2 and 2.4 GB over 40 consecutive compiles and ended the full benchmark at 1.8 GB.
   The reason a compile allocates that much native memory in Revit (about 400 MB, against 12 MB managed) is still not understood.
2. **IronPython pays for cancellation.** Stopping runaway loops uses a per-line trace hook. Measured headless: the hook makes loops 1.7x slower and
   recursion 6x slower (fib(27) 300 ms against 52 ms without it). Debug mode itself is free. Options: make tracing opt-in per job, or cancel through a
   cheaper mechanism. Not changed yet; the trade-off is safety against speed.
3. **Roslyn compiles every job from scratch (about 300 to 450 ms).** The same compile is about 90 ms outside Revit. A cache keyed by the script text would remove
   it for repeated scripts; reusing sessions already avoids it for follow-up calls.
4. **`value.text` of a collection is only its .NET type name** (for example `System.Collections.Generic.List`1[System.Int32]`); the data is in `value.json`.

## Reproducing

```
tests/bench/cold.ps1   -Port 23717 -Engine roslyn          # fresh Revit each time
tests/bench/open-model.ps1 -Port 23717 -Path "<model.rvt>"
tests/bench/bench.ps1  -Port 23717 -Tasks tests/bench/tasks -Out bench.json -Runs 7
```

Single machine, single model, 7 runs per cell: treat differences below roughly 20% as noise.
