using System;
using NCalc;

namespace Курсовой_проект.Services
{
    public class ExpressionParser
    {
        private string _expression;

        public ExpressionParser(string expression)
        {
            _expression = expression;
        }

        public Func<double, double, double> ParseFunction()
        {
            return (x, y) =>
            {
                try
                {
                    var expression = new Expression(_expression, EvaluateOptions.IgnoreCase);

                    expression.Parameters["x"] = x;
                    expression.Parameters["y"] = y;

                    object result = expression.Evaluate();

                    // Преобразуем результат в double
                    if (result is double)
                        return (double)result;
                    if (result is int)
                        return (int)result;
                    if (result is decimal)
                        return (double)(decimal)result;

                    return Convert.ToDouble(result);
                }
                catch (Exception ex)
                {
                    throw new Exception($"Ошибка вычисления f({x},{y}): {ex.Message}");
                }
            };
        }

        public bool ValidateExpression(out string errorMessage)
        {
            errorMessage = string.Empty;
            try
            {
                var expression = new Expression(_expression, EvaluateOptions.IgnoreCase);
                expression.Parameters["x"] = 0.5;
                expression.Parameters["y"] = 0.5;
                expression.Evaluate();
                return true;
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                return false;
            }
        }
    }
}