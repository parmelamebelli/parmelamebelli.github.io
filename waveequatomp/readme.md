# Parallel Wave Equation Simulation with OpenMP

## Overview:
This program simulates the wave equation using OpenMP to exploit parallel processing capabilities, optimizing the solution over a given spatial-temporal grid. The solution provides both initial conditions and boundary conditions for the wave propagation over the domain.

## Wave Equation:
The classical wave equation in one dimension is given by:

    ∂²u/∂t² = c² ∂²u/∂x²

where:
- `u` is the amplitude of the wave at position `x` and time `t`.
- `c` is the speed of propagation of the wave.

In this program, the wave equation is discretized using finite differences. The program focuses on the time-stepping solution for the wave, with given initial and boundary conditions.

## Parallel Programming with OpenMP:
The code has been enhanced to use OpenMP, a parallel programming API, to accelerate the computation. The primary parallelized segments are the loops iterating over the spatial grid points. By splitting the iterations among available processor cores, the simulation can complete more rapidly, especially when `m` and `n` are large.



## Code Parameters:

### Constants:
- `L` (double): Length of the spatial domain. Set to `1.0`.
- `T` (double): Total time for the simulation. Set to `0.5`.
- `m` (const int): Number of discretization points in the spatial domain. Set to `4`.
- `n` (const int): Number of discretization points in the temporal domain. Set to `4`.

### Derived Parameters:
- `h` (double): Spatial discretization step. Computed as `L / m`.
- `k` (double): Temporal discretization step. Computed as `T / n`.
- `r` (double): Ratio of the square of temporal to spatial discretization. Computed as `(k/h)^2`.
