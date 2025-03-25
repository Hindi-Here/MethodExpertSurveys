using System;
using System.IO;
using System.Linq;

using OfficeOpenXml;

namespace MethodExpertSurveys.Views
{
    internal static class CalculateRanking
    {
        public static double[,] CreateMatrix(string filePath, string range)
        {
            using ExcelPackage package = new(new FileInfo(filePath));
            var worksheet = package.Workbook.Worksheets[0];
            var cells = worksheet.Cells[range];

            int row_start = cells.Start.Row;
            int column_start = cells.Start.Column;
            int row_end = cells.End.Row;
            int column_end = cells.End.Column;

            int rows = row_end - row_start + 1;
            int cols = column_end - column_start + 1;
            double[,] matrix = new double[rows, cols];

            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    matrix[i, j] = Convert.ToDouble(cells[row_start + i, column_start + j].Value);
                }
            }
            return matrix;
        }

        public static int[] DirectRanking(double[,] matrix)
        {
            double[] answerScore = GetAnswerScore(matrix);
            double[] weightingFactor = GetWeightingFactor(answerScore);
            return GetRanks(weightingFactor);
        }

        public static double[] GetAnswerScore(double[,] matrix)
        {
            int rows = matrix.GetLength(0);
            int cols = matrix.GetLength(1);
            double[] answerScore = new double[rows];
            for (int i = 0; i < rows; i++)
            {
                double sum = 0;
                for (int j = 0; j < cols; j++)
                {
                    sum += matrix[i, j];
                }
                answerScore[i] = sum;
            }
            return answerScore;
        }

        public static double[] GetWeightingFactor(double[] answerScore)
        {
            double totalSum = answerScore.Sum();
            double[] weightingFactor = new double[answerScore.Length];

            for (int i = 0; i < answerScore.Length; i++)
                weightingFactor[i] = Math.Round(answerScore[i] / totalSum, 5);

            return weightingFactor;
        }

        public static int[] GetRanks(double[] weightingFactor)
        {
            int length = weightingFactor.Length;
            int[] ranks = new int[length];

            var sorted = Enumerable.Range(0, length)
                                   .OrderBy(i => weightingFactor[i])
                                   .ToArray();

            int rank = 1;
            for (int i = 0; i < sorted.Length; i++)
            {
                if (i > 0 && Math.Abs(weightingFactor[sorted[i]] - weightingFactor[sorted[i - 1]]) < double.Epsilon)
                {
                    ranks[sorted[i]] = ranks[sorted[i - 1]];
                }
                else
                {
                    ranks[sorted[i]] = rank;
                    rank++;
                }
            }

            return ranks;
        }



        public static int[] PairedComparison(double[,] matrix)
        {
            double[,] newMatrix = CreateNewMatrix(matrix);
            return DirectRanking(newMatrix);
        }

        public static int[] GetDominanceRanks(int[] dominanceCount)
        {
            int length = dominanceCount.Length;
            int[] ranks = new int[length];

            var sorted = Enumerable.Range(0, length)
                                   .OrderBy(i => dominanceCount[i])
                                   .ToArray();

            for (int i = 0; i < length; i++)
            {
                ranks[sorted[length - 1 - i]] = i + 1;
            }

            return ranks;
        }

        public static int[] GetDominanceCount(double[,] matrix, int column_number)
        {
            int rows = matrix.GetLength(0);
            int[] dominanceCount = new int[rows];
            for (int i = 0; i < rows; i++)
            {
                int count = 0;
                for (int j = 0; j < rows; j++)
                {
                    if (matrix[i, column_number] < matrix[j, column_number])
                    {
                        count++;
                    }
                }
                dominanceCount[i] = count;
            }

            return dominanceCount;
        }

        public static double[] GetExpertWeightingFactor(double[,] matrix, int column_number)
        {
            double totalSum = 0;
            for (int i = 0; i < matrix.GetLength(0); i++)
                totalSum += i;

            int[] dominanceCount = GetDominanceCount(matrix, column_number);

            double[] weightingFactor = new double[dominanceCount.Length];
            for (int i = 0; i < dominanceCount.Length; i++)
                weightingFactor[i] = Math.Round(dominanceCount[i] / totalSum, 5);

            return weightingFactor;
        }

        public static double[,] CreateNewMatrix(double[,] matrix)
        {
            int rows = matrix.GetLength(0);
            int cols = matrix.GetLength(1);
            double[,] newMatrix = new double[rows, cols];
            for (int j = 0; j < cols; j++)
            {
                int[] ranks = GetDominanceRanks(GetDominanceCount(matrix, j));

                for (int i = 0; i < rows; i++)
                {
                    newMatrix[i, j] = ranks[i];
                }
            }

            return newMatrix;
        }

    }
}
