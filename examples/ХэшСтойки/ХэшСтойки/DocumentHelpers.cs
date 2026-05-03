using System;
using TFlex.Model;

namespace TFAPIHash
{
    public static class vb
    {
        public static int GetIntVariableValue(Document document, string VariableName)
        {
            Variable var = document.FindVariable(VariableName);
            if (var == null) { throw new Exception($"Переменная {VariableName} не найдена"); }
            return ((int)var.RealValue);
        }
        public static double GetDoubleVariableValue(Document document, string VariableName)
        {
            Variable var = document.FindVariable(VariableName);
            if (var == null) { throw new Exception($"Переменная {VariableName} не найдена"); }
            return (var.RealValue);
        }

        public static string GetStringVariableValue(Document document, string VariableName)
        {
            Variable var = document.FindVariable(VariableName);
            if (var == null) { throw new Exception($"Переменная {VariableName} не найдена"); }
            return (var.TextValue);
        }

    }
}
