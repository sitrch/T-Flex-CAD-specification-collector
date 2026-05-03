# TFlexAPI.dll - Исследование происхождения записей спецификации

## Определение происхождения записи спецификации

Задача: определить, наследуется ли запись спецификации от дочернего элемента, или создаётся в текущем фрагменте.

## Структура классов

### BOMRecord (TFlex.Model\BOMRecord.cs)
- Запись данных для спецификации
- Содержит внутреннее поле `m_rowElem` типа `RowElement`

### RowElement (TFlex.Model\RowElement.cs)
- Элемент структуры изделия, на котором базируется запись спецификации
- Содержит свойства для определения происхождения записи

## Свойства для определения происхождения

### Из RowElement (TFlex.Model)

| Свойство | Тип | Описание |
|---------|-----|---------|
| `SourceFragmentFirstLevel` | `Fragment` | Фрагмент, из которого поднят элемент. `null`, если не поднят из фрагмента |
| `SourceFragment3DFirstLevel` | `ModelObject` | Фрагмент 3D, из которого поднят элемент. `null`, если не поднят из 3D фрагмента |
| `SourceFragmentPath` | `string` | Путь к фрагменту, из которого поднят элемент. `null`, если не поднят |
| `SourceRowElementUIDFirstLevel` | `Guid` | UID элемента первого уровня, по которому создан этот элемент. `Guid.Empty`, если не поднят из фрагмента |
| `SourceRowElementUID` | `Guid` | UID элемента, по которому создан этот элемент. `Guid.Empty`, если не поднят из фрагмента |
| `SourceObject` | `ModelObject` | Исходный модельный объект, по которому создана запись. `null`, если элемент не собран по текущему документу |

### Из Element (TFlex.Model.Structure)

| Свойство | Тип | Описание |
|---------|-----|---------|
| `SourceFragment` | `Fragment` | Фрагмент, из которого поднят элемент. `null`, если не поднят |
| `SourceFragmentUID` | `Guid` | UID фрагмента источника |
| `External` | `ExternalMode` | Режим внешнего элемента (None, OneLevel, AllLevels) |

## Критерии определения

### Запись НАСЛЕДУЕТСЯ от дочернего элемента, если:
- `rowElement.SourceFragmentFirstLevel != null`
- ИЛИ `rowElement.SourceRowElementUID != Guid.Empty`
- ИЛИ `rowElement.SourceObject != null`

### Запись СОЗДАНА в текущем фрагменте, если:
- `rowElement.SourceFragmentFirstLevel == null`
- И `rowElement.SourceRowElementUID == Guid.Empty`
- И `rowElement.SourceObject == null`

## Пример использования

```csharp
public bool IsRecordInheritedFromChildElement(BOMRecord record)
{
    if (record == null) return false;
    
    // Через рефлексию получаем доступ к m_rowElem
    // или используем доступные методы/свойства BOMRecord
    
    bool hasSourceFragment = (record.m_rowElem.SourceFragmentFirstLevel != null);
    bool hasSourceUID = (record.m_rowElem.SourceRowElementUID != Guid.Empty);
    
    return hasSourceFragment || hasSourceUID;
}

public bool IsRecordCreatedInCurrentFragment(BOMRecord record)
{
    if (record == null) return false;
    
    return (record.m_rowElem.SourceFragmentFirstLevel == null && 
            record.m_rowElem.SourceRowElementUID == Guid.Empty);
}
```

## Вывод

Основной индикатор происхождения записи — свойство `SourceFragmentFirstLevel` класса `RowElement`:
- `null` = запись создана локально в текущем документе
- Объект `Fragment` = запись наследуется из дочернего фрагмента

## Дополнительные методы
### StructureElementsManager содержит метод для получения элементов, загруженных из фрагмента:
```csharp
List<Element> GetElementsLoadedFromFragment(Fragment fragment)
```

### Проверка цепочки фрагментов:
```csharp
IEnumerable<ObjectId> GetSourceFragmentIdChain(bool fragment3d)
```

---

## Получение флага включения в спецификацию

### Из BOMRecord (прямой способ)
```csharp
// Для спецификации сборочного документа
bool includeInAssembly = bomRecord.IncludeToAssemblyBOM;

// Для спецификации текущего документа  
bool includeInCurrentDoc = bomRecord.IncludeToCurrentDocumentBOM;
```
**Источник**: `T_TFlex_Model_BOMRecord.htm` — свойства:
- `IncludeToAssemblyBOM` (bool) — "Включать запись в спецификацию сборочного документа"
- `IncludeToCurrentDocumentBOM` (bool) — "Включать запись в спецификацию текущего документа"

### Через RowElement (внутренний доступ)
`BOMRecord` содержит внутреннее поле `m_rowElem` типа `RowElement`:

```csharp
// Доступ через рефлексию
var field = typeof(BOMRecord).GetField("m_rowElem", 
    System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
var rowElement = field.GetValue(bomRecord) as RowElement;

if (rowElement != null)
{
    RowElementCell includeCell = rowElement.IncludeInDoc; // или IncludeInAssembly
    bool includeValue = Convert.ToBoolean(includeCell.Value);
}
```

### Свойства RowElement для включения
- `IncludeInAssembly` (RowElementCell) — "Получить ячейку 'Включать в отчёты/спецификации текущего документа'"
- `IncludeInDoc` (RowElementCell) — аналогично

### Связь с настройками фрагмента
Если нужно узнать, **почему** стоит флаг:
- `Fragment2D.NeverIncludeInPS` (bool) — для PS (Product Structure)
- `Fragment3D.IncludeInBOM` (`IncludeInBOMType`) — для BOM

### IncludeInBOMType (перечисление)
| Значение | Описание |
|-----------|-----------|
| `None` (0) | Не включать |
| `Default` (1) | По умолчанию |
| `WithoutEmbeddedElems` (2) | Без вложенных элементов |
| `WithEmbeddedElems` (3) | С вложенными элементами |
| `OnlyEmbeddedElems` (4) | Только вложенные элементы |

### UseStatusType (перечисление для статуса по умолчанию)
| Значение | Описание |
|-----------|-----------|
| `Default` (0) | По умолчанию |
| `FragmentDocument` (1) | Использовать документ фрагмента |
| `AssemblyDocument` (2) | Использовать документ сборки |

**Вывод**: Для получения значения флага включения в спецификацию **используйте свойства `BOMRecord.IncludeToAssemblyBOM` или `IncludeToCurrentDocumentBOM`**.

**Проверка цепочки фрагментов:**
```csharp
IEnumerable<ObjectId> GetSourceFragmentIdChain(bool fragment3d)
```
