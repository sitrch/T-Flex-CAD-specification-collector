using System;
using System.Collections.Generic;

using TFlex.CadReportGenerator;

namespace Gen
{
    public class GenRows
    {
        private const string BeforeAttrName = "EmptyRows_CountBefore"; // Имя атрибута отчета. Число пустых строк перед каждым элементом.
        private const string AfterAttrName = "EmptyRows_CountAfter";// Число пустых строк после каждого элемента.
        private const string SkipFirstAttrName = "EmptyRows_SkipFirst";// Пропускать первую
        private const string SkipLastAttrName = "EmptyRows_SkipLast";// Пропускать последнюю


        public class DummyItemInfo : ItemInfo
        {
            private static HashSet<String> AllowedParameters = new HashSet<string> { "раздел" };

            public DummyItemInfo(ItemInfo mainItem) : base(mainItem.PSInfo) { MainItemInfo = mainItem; }
            public ItemInfo MainItemInfo { get; private set; }
            public String Value { get; set; }
            public override string this[string paramName]
            {
                get
                {
                    string invParamName = paramName.ToLowerInvariant();
                    if (AllowedParameters.Contains(invParamName))
                        return MainItemInfo[paramName];
                    return String.Empty;
                }
            }

            public override object GetValue(string paramName)
            {
                string invParamName = paramName.ToLowerInvariant();
                if (AllowedParameters.Contains(invParamName))
                    return MainItemInfo.GetValue(invParamName);
                return String.Empty;
            }
        }


        public static bool InsertEmptyRowsMacro(MacroCallContext context, GroupItemInfo group)
        {
            var generAttrs = context.Properties.Attributes;
            int before = generAttrs.HasAttribute(BeforeAttrName) ? generAttrs[BeforeAttrName].ValueAsInt : 0;
            int after = generAttrs.HasAttribute(AfterAttrName) ? generAttrs[AfterAttrName].ValueAsInt : 0;
            if (before <= 0 && after <= 0)
                return true;

            bool skipFirst = generAttrs.HasAttribute(SkipFirstAttrName) ? generAttrs[SkipFirstAttrName].ValueAsBool : false;
            bool skipLast = generAttrs.HasAttribute(SkipLastAttrName) ? generAttrs[SkipLastAttrName].ValueAsBool : false;

            List<ItemInfo> newChildren = new List<ItemInfo>();
            for (int nIndex = 0; nIndex < group.Children.Count; ++nIndex)
            {
                var item = group.Children[nIndex];
                if (item is GroupItemInfo)
                    newChildren.Add(item);
                else
                {
                    if (!(skipFirst && nIndex == 0))
                        for (int i = 0; i < before; ++i)
                            newChildren.Add(new DummyItemInfo(item));
                    newChildren.Add(item);
                    if (!(skipLast && nIndex == (group.Children.Count - 1)))
                        for (int i = 0; i < after; ++i)
                            newChildren.Add(new DummyItemInfo(item));
                }
            }
            group.Children = newChildren;
            return true;
        }
    }
}






