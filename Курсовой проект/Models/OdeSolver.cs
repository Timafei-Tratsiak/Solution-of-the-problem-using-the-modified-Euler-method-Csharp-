using System;
using System.Collections.Generic;

namespace Курсовой_проект.Models
{
    public class OdeSolver
    {
        public List<ResultPoint> SolveModifiedEuler(
            Func<double, double, double> f,
            double x0,
            double y0,
            double X,
            double h)
        {
            var results = new List<ResultPoint>();

            double x = x0;
            double y = y0;

            results.Add(new ResultPoint(x, y, f(x, y)));

            int steps = (int)Math.Round((X - x0) / h);

            for (int i = 0; i < steps; i++)
            {
                double new_y = y + h * f(x + h / 2, y + (h / 2) * f(x, y));
                x += h;
                y = new_y;
                results.Add(new ResultPoint(x, y, f(x, y)));
            }

            return results;
        }
    }
}