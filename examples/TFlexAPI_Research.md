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

**StructureElementsManager** содержит метод для получения элементов, загруженных из фрагмента:
```csharp
List<Element> GetElementsLoadedFromFragment(Fragment fragment)
```

**Проверка цепочки фрагментов:**
```csharp
IEnumerable<ObjectId> GetSourceFragmentIdChain(bool fragment3d)
```
