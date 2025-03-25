using OfficeOpenXml.Style;
using OfficeOpenXml;

using System.Collections.Generic;
using System.IO;

namespace MethodExpertSurveys.Views
{
    internal static class RankingBuilder
    {
        static RankingBuilder() => ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        public static void DirectRankingBuild(string inputPath, string range, string outputPath)
        {
            using ExcelPackage package = new();
            var worksheet = package.Workbook.Worksheets.Add("Experts");

            double[,] matrix = CalculateRanking.CreateMatrix(inputPath, range);

            WriteStartMatrix(worksheet, matrix);
            WriteCalculatePart(worksheet, matrix);
            SetTableBorders(worksheet, matrix);

            for (int i = 1; i <= matrix.GetLength(1) + 5; i++)
                worksheet.Column(i).AutoFit();

            FileInfo fileInfo = new(outputPath + "\\DirRanking.xlsx");
            package.SaveAs(fileInfo);
        }

        private static void WriteStartMatrix(ExcelWorksheet worksheet, double[,] matrix)
        {
            for (int i = 0; i < matrix.GetLength(1); i++)
            {
                worksheet.Cells[2, i + 3].Value = $"Эксперт {i + 1}";
                worksheet.Cells[2, i + 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[2, i + 3].Style.Font.Bold = true;
            }

            for (int i = 0; i < matrix.GetLength(0); i++)
            {
                worksheet.Cells[i + 3, 2].Value = $"Оценка {i + 1}";
                worksheet.Cells[i + 3, 2].Style.Font.Bold = true;
            }

            for (int i = 0; i < matrix.GetLength(0); i++)
            {
                for (int j = 0; j < matrix.GetLength(1); j++)
                {
                    worksheet.Cells[i + 3, j + 3].Value = matrix[i, j];
                    worksheet.Cells[i + 3, j + 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                }
            }
        }

        private static void WriteCalculatePart(ExcelWorksheet worksheet, double[,] matrix)
        {
            int column = matrix.GetLength(1) + 3;

            worksheet.Cells[2, column].Value = "Сумма";
            worksheet.Cells[2, column + 1].Value = "Вес";
            worksheet.Cells[2, column + 2].Value = "Ранг";
            for (int i = 0; i < 3; i++)
            {
                worksheet.Cells[2, column + i].Style.Font.Bold = true;
                worksheet.Cells[2, column + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }

            double[] answerScore = CalculateRanking.GetAnswerScore(matrix);
            for (int i = 0; i < answerScore.Length; i++)
            {
                worksheet.Cells[i + 3, column].Value = answerScore[i];
                worksheet.Cells[i + 3, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }

            double[] weightFactor = CalculateRanking.GetWeightingFactor(answerScore);
            for (int i = 0; i < weightFactor.Length; i++)
            {
                worksheet.Cells[i + 3, column + 1].Value = weightFactor[i];
                worksheet.Cells[i + 3, column + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }

            int[] rank = CalculateRanking.GetRanks(weightFactor);
            for (int i = 0; i < rank.Length; i++)
            {
                worksheet.Cells[i + 3, column + 2].Value = rank[i];
                worksheet.Cells[i + 3, column + 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
        }

        private static void SetTableBorders(ExcelWorksheet worksheet, double[,] matrix)
        {
            int row_start = 2;
            int column_start = 2;
            int row_end = matrix.GetLength(0) + row_start;
            int column_end = matrix.GetLength(1) + column_start + 3;

            for (int i = row_start; i <= row_end; i++)
            {
                worksheet.Cells[i, column_start].Style.Border.Left.Style = ExcelBorderStyle.Medium;
            }

            for (int i = row_start; i <= row_end; i++)
            {
                worksheet.Cells[i, column_end].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                worksheet.Cells[i, column_end - 3].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            }

            for (int i = column_start; i <= column_end; i++)
            {
                worksheet.Cells[row_start, i].Style.Border.Top.Style = ExcelBorderStyle.Medium;
            }

            for (int i = column_start; i <= column_end; i++)
            {
                worksheet.Cells[row_end, i].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
            }
        }



        public static void PairComRankingBuild(string inputPath, string range, string outputPath)
        {
            using ExcelPackage package = new();
            var worksheet = package.Workbook.Worksheets.Add("Experts");

            double[,] temp_matrix = CalculateRanking.CreateMatrix(inputPath, range);
            double[,] matrix = CalculateRanking.CreateNewMatrix(temp_matrix);

            WriteStartMatrix(worksheet, matrix);
            WriteCalculatePart(worksheet, matrix);
            SetTableBorders(worksheet, matrix);

            for (int i = 1; i <= matrix.GetLength(1) + 5; i++)
                worksheet.Column(i).AutoFit();

            var expertsSheet = package.Workbook.Worksheets.Add("Individual  Experts");
            WriteAllIndividualTables(expertsSheet, matrix);

            FileInfo fileInfo = new(outputPath + "\\PairComRanking.xlsx");
            package.SaveAs(fileInfo);
        }

        private static void WriteAllIndividualTables(ExcelWorksheet worksheet, double[,] matrix)
        {
            int row = 1;
            int expert_cols = matrix.GetLength(1);

            for (int i = 0; i < expert_cols; i++)
            {
                row = WriteIndividualTable(worksheet, matrix, i, row);
            }

            for (int col = 1; col <= 7; col++)
            {
                worksheet.Column(col).AutoFit();
            }
        }

        private static int WriteIndividualTable(ExcelWorksheet worksheet, double[,] matrix, int expertIndex, int startRow)
        {
            int rows = matrix.GetLength(0);
            int dataRow = startRow + 1;
            int endRow = dataRow + rows;

            WriteIndividualTableHeader(worksheet, expertIndex, startRow);
            WriteIndividualTableBody(worksheet, matrix, expertIndex, dataRow, rows);
            SetIndividualTableBorder(worksheet, dataRow, endRow);

            return endRow + 2;
        }

        private static void WriteIndividualTableHeader(ExcelWorksheet worksheet, int expertIndex, int row)
        {
            worksheet.Cells[row, 2].Value = $"ЭКСПЕРТ {expertIndex + 1}";
            worksheet.Cells[row, 2].Style.Font.Bold = true;
            worksheet.Cells[row, 2].Style.Font.Size = 12;

            string[] headers = ["Параметр", "Ранг", "Доминирует над параметрами:", "Сумма", "Вес"];
            for (int i = 0; i < headers.Length; i++)
            {
                worksheet.Cells[row + 1, 2 + i].Value = headers[i];
                worksheet.Cells[row + 1, 2 + i].Style.Font.Bold = true;
                worksheet.Cells[row + 1, 2 + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
        }

        private static void WriteIndividualTableBody(ExcelWorksheet worksheet, double[,] matrix, int expertIndex, int dataRow, int rows)
        {
            int[] dominanceCount = CalculateRanking.GetDominanceCount(matrix, expertIndex);
            double[] expertWeights = CalculateRanking.GetExpertWeightingFactor(matrix, expertIndex);

            for (int i = 0; i < rows; i++)
            {
                int row = dataRow + 1 + i;

                worksheet.Cells[row, 2].Value = i + 1;
                worksheet.Cells[row, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells[row, 3].Value = matrix[i, expertIndex];
                worksheet.Cells[row, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                var dominateString = new List<string>();
                for (int j = 0; j < rows; j++)
                {
                    if (matrix[i, expertIndex] < matrix[j, expertIndex])
                    {
                        dominateString.Add((j + 1).ToString());
                    }
                }
                worksheet.Cells[row, 4].Value = string.Join(", ", dominateString);

                worksheet.Cells[row, 5].Value = dominanceCount[i];
                worksheet.Cells[row, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells[row, 6].Value = expertWeights[i];
                worksheet.Cells[row, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[row, 6].Style.Numberformat.Format = "0.00000";
            }
        }

        private static void SetIndividualTableBorder(ExcelWorksheet worksheet, int dataRow, int endRow)
        {
            worksheet.Cells[dataRow, 2, dataRow, 6].Style.Border.Top.Style = ExcelBorderStyle.Medium;
            worksheet.Cells[endRow, 2, endRow, 6].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
            worksheet.Cells[dataRow, 2, endRow, 2].Style.Border.Left.Style = ExcelBorderStyle.Medium;
            worksheet.Cells[dataRow, 6, endRow, 6].Style.Border.Right.Style = ExcelBorderStyle.Medium;
        }

    }
}
