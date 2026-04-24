namespace Курсовой_проект.Models
{
    public class ResultPoint
    {
        public double X { get; set; }
        public double Y { get; set; }
        public double F { get; set; }

        public ResultPoint(double x, double y, double f)
        {
            X = x;
            Y = y;
            F = f;
        }
    }
}