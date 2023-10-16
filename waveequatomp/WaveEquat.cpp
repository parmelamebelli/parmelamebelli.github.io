#include <omp.h>
#include <cstdio>
#include <iostream> 
#include <cmath>

using namespace std;

int main() 
{ 
    double fillimi = omp_get_wtime();
    const int m = 4, n = 4;
    const double L = 1.0, T = 0.5;

    double h = L / m;
    double k = T / n;
    double r = std::pow((k / h), 2);

    double x[m + 1], y[n + 1];
    double u[m + 1][n + 1], v[m + 1][n + 1];

    #pragma omp parallel for
    for (int i = 0; i <= m; i++) 
        x[i] = i * h;

    #pragma omp parallel for
    for (int j = 0; j <= n; j++) 
        y[j] = j * k;

    // Initial conditions
    #pragma omp parallel for
    for (int i = 1; i < m; i++) 
    {
        u[i][0] = x[i] + 2;
        v[i][0] = x[i];
    }

    // Boundary conditions
    for (int j = 0; j <= n; j++) 
    {
        u[0][j] = 0;
        u[m][j] = 2 * y[j];
    }

    #pragma omp parallel for
    for (int i = 1; i <= m - 1; i++) 
        u[i][1] = r * u[i - 1][0] + (1 - 2 * r) * u[i][0] + r * u[i + 1][0] + k * v[i][0];

    for (int j = 1; j < n; j++) 
    {
        #pragma omp parallel for
        for (int i = 1; i < m; i++) 
            u[i][j + 1] = r * u[i - 1][j] + 2 * (1 - r) * u[i][j] + r * u[i + 1][j] - u[i][j - 1];
    }

    for (int i = 0; i < m + 1; i++) 
    {
        for (int j = 0; j < n + 1; j++) 
            std::printf("u[%d][%d] = %.16g \n", i, j, u[i][j]);
    }

    double ke = omp_get_wtime() - fillimi;
    printf("Koha e ekzekutimit = %f \n", ke);

    return 0;
}
