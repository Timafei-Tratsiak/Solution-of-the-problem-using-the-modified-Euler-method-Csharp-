using System;
using System.Windows;
using MLApp;

namespace Курсовой_проект.Services
{
    public class MatlabIntegrationService
    {
        private MLApp.MLApp _matlab; // Сохраняем экземпляр, чтобы MATLAB не закрывался

        /// <summary>
        /// Запускает MATLAB и решает ДУ методом ode45 (Рунге-Кутта 4/5 порядка)
        /// </summary>
        public void SolveWithOde45(double x0, double y0, double xEnd, double h, string function)
        {
            try
            {
                // Создаём COM-объект MATLAB только один раз
                if (_matlab == null)
                {
                    _matlab = new MLApp.MLApp();
                    _matlab.Visible = 1; // Показываем окно MATLAB
                }

                // Передаём данные в рабочее пространство MATLAB
                _matlab.PutWorkspaceData("x0", "base", x0);
                _matlab.PutWorkspaceData("y0", "base", y0);
                _matlab.PutWorkspaceData("x_end", "base", xEnd);
                _matlab.PutWorkspaceData("func_str", "base", function);

                // Выполняем скрипт с ode45 (английское название для файла)
                string script = @"
                    % Преобразуем строку в анонимную функцию
                    f = str2func(['@(x,y)' func_str]);
                    
                    % Определяем функцию для ode45
                    dydx = @(x,y) f(x, y);
                    
                    % Решаем ДУ методом ode45 (адаптивный Рунге-Кутта 4/5 порядка)
                    [x_vals, y_vals] = ode45(dydx, [x0, x_end], y0);
                    
                    % Строим график решения (название на английском)
                    figure('Name', 'Solution of ODE by Modified Euler Method', 'NumberTitle', 'off', 'Position', [100, 100, 800, 600]);
                    plot(x_vals, y_vals, 'b-o', 'LineWidth', 2, 'MarkerSize', 4);
                    grid on;
                    xlabel('x');
                    ylabel('y(x)');
                    title(sprintf('Solution of ODE by Modified Euler Method (Runge-Kutta 4th order)\n dy/dx = %s', func_str));
                    legend('Numerical solution (ode45)');
                    
                    % Выводим информацию в консоль MATLAB
                    fprintf('=== Solution by ode45 method ===\n');
                    fprintf('Initial conditions: x0 = %.4f, y0 = %.4f\n', x0, y0);
                    fprintf('Interval: [%.4f, %.4f]\n', x0, x_end);
                    fprintf('Number of solution points: %d\n', length(x_vals));
                    fprintf('Final value: y(%.4f) = %.6f\n', x_end, y_vals(end));
                    fprintf('=================================\n');
                    
                    % Сохраняем график в файл (английское название)
                    saveas(gcf, 'ODE_Solution_by_Modified_Euler_Method.png');
                    disp('Graph saved as: ODE_Solution_by_Modified_Euler_Method.png');
                    disp('Matlab window will remain open. Close it manually when done.');
                ";

                _matlab.Execute(script);

                // Не вызываем _matlab.Quit() — MATLAB остаётся открытым
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error calling MATLAB ode45: {ex.Message}\n\n" +
                    "Please ensure:\n" +
                    "1. MATLAB R2016a is installed\n" +
                    "2. COM reference to MATLAB Application Type Library is added\n" +
                    "3. Project is built for x64 platform\n" +
                    "4. MATLAB is registered as COM server",
                    "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}