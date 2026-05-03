#region сборка TFlexAPI, Version=17.1.36.0, Culture=neutral, PublicKeyToken=eab6a180a6be0d77
// C:\Program Files\T-FLEX CAD 17\Program\TFlexAPI.dll
// Decompiled with ICSharpCode.Decompiler 9.1.0.7988
#endregion

using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using AssemblyInfo;
using ATL;
using ps;
using std;
using TFlex.Model.Data.ProductStructure;
using TFlex.Model.Model2D;
using TFlex.Model.Units;

namespace TFlex.Model;

//
// Сводка:
//     Класс элемента структуры изделия
public class RowElement : ModelObject
{
    //
    // Сводка:
    //     Получить ячейку "Включать при вставке в сборку"
    public unsafe RowElementCell Position
    {
        get
        {
            RowElementCell result = new RowElementCell(this, *global::_003CModule_003E.ps_002EParameterDescriptor_002EGetID(global::_003CModule_003E.ps_002EScheme_002EGetAuxiliaryParameter((ps.Scheme.AuxiliaryParam)2)));
            GC.KeepAlive(this);
            return result;
        }
    }

    //
    // Сводка:
    //     Получить ячейку "Включать в отчёты/спецификации текущего документа"
    public unsafe RowElementCell IncludeInAssembly
    {
        get
        {
            RowElementCell result = new RowElementCell(this, *global::_003CModule_003E.ps_002EParameterDescriptor_002EGetID(global::_003CModule_003E.ps_002EScheme_002EGetAuxiliaryParameter((ps.Scheme.AuxiliaryParam)1)));
            GC.KeepAlive(this);
            return result;
        }
    }

    //
    // Сводка:
    //     Получить ячейку "Включать в отчёты/спецификации текущего документа"
    public unsafe RowElementCell IncludeInDoc
    {
        get
        {
            RowElementCell result = new RowElementCell(this, *global::_003CModule_003E.ps_002EParameterDescriptor_002EGetID(global::_003CModule_003E.ps_002EScheme_002EGetAuxiliaryParameter((ps.Scheme.AuxiliaryParam)0)));
            GC.KeepAlive(this);
            return result;
        }
    }

    public RowElementCell this[Guid parameterId] => GetCell(parameterId);

    public RowElementCell this[TFlex.Model.Data.ProductStructure.ParameterDescriptor parameter] => GetCell(parameter);

    //
    // Сводка:
    //     Модельные объекты, связанные с элементом структуры изделия
    public RowElementLinkedObjects LinkedObjects => new RowElementLinkedObjects(this);

    //
    // Сводка:
    //     Уникальный идентификатор элемента, по которому создан этот элемент (первого уровня
    //     вложенности). Если элемент не поднят из фрагмента, то возвращается System.Guid.Empty
    public unsafe Guid SourceRowElementUIDFirstLevel
    {
        get
        {
            //IL_004f: Expected I, but got I8
            AssemblyRowElement* ptr = (AssemblyRowElement*)global::_003CModule_003E.__RTDynamicCast(base.repGet, 0, System.Runtime.CompilerServices.Unsafe.AsPointer(ref global::_003CModule_003E._003F_003F_R0_003FAVRowElement_0040ps_0040_0040_00408), System.Runtime.CompilerServices.Unsafe.AsPointer(ref global::_003CModule_003E._003F_003F_R0_003FAVAssemblyRowElement_0040ps_0040_0040_00408), 0);
            if (ptr != null)
            {
                FragRowElemID* ptr2 = global::_003CModule_003E.ps_002EAssemblyRowElement_002EGetFragRowElemID(ptr);
                if (global::_003CModule_003E.ps_002EFragRowElemID_002EIsOneLevelID(ptr2))
                {
                    UniqueIdentifier uniqueIdentifier = *(UniqueIdentifier*)ptr2;
                    Guid result = global::_003CModule_003E.TFlex_002EGUID2Guid(global::_003CModule_003E.UniqueIdentifier_002E_002EAEBU_GUID_0040_0040(&uniqueIdentifier));
                    GC.KeepAlive(this);
                    return result;
                }

                if (global::_003CModule_003E.std_002Evector_003Cps_003A_003AFragID_002Cstd_003A_003Aallocator_003Cps_003A_003AFragID_003E_0020_003E_002Esize((vector_003Cps_003A_003AFragID_002Cstd_003A_003Aallocator_003Cps_003A_003AFragID_003E_0020_003E*)((ulong)(nint)ptr2 + 32uL)) == 1)
                {
                    UniqueIdentifier uniqueIdentifier2 = *(UniqueIdentifier*)((ulong)(nint)ptr2 + 16uL);
                    Guid result2 = global::_003CModule_003E.TFlex_002EGUID2Guid(global::_003CModule_003E.UniqueIdentifier_002E_002EAEBU_GUID_0040_0040(&uniqueIdentifier2));
                    GC.KeepAlive(this);
                    return result2;
                }
            }

            GC.KeepAlive(this);
            return Guid.Empty;
        }
    }

    //
    // Сводка:
    //     Уникальный идентификатор элемента, по которому создан этот элемент. Если элемент
    //     не поднят из фрагмента, то возвращается System.Guid.Empty
    public unsafe Guid SourceRowElementUID
    {
        get
        {
            AssemblyRowElement* ptr = (AssemblyRowElement*)global::_003CModule_003E.__RTDynamicCast(base.repGet, 0, System.Runtime.CompilerServices.Unsafe.AsPointer(ref global::_003CModule_003E._003F_003F_R0_003FAVRowElement_0040ps_0040_0040_00408), System.Runtime.CompilerServices.Unsafe.AsPointer(ref global::_003CModule_003E._003F_003F_R0_003FAVAssemblyRowElement_0040ps_0040_0040_00408), 0);
            if (ptr != null)
            {
                FragRowElemID* ptr2 = global::_003CModule_003E.ps_002EAssemblyRowElement_002EGetFragRowElemID(ptr);
                UniqueIdentifier uniqueIdentifier = *(UniqueIdentifier*)((ulong)(nint)ptr2 + 16uL);
                Guid result = global::_003CModule_003E.TFlex_002EGUID2Guid(global::_003CModule_003E.UniqueIdentifier_002E_002EAEBU_GUID_0040_0040(&uniqueIdentifier));
                GC.KeepAlive(this);
                return result;
            }

            GC.KeepAlive(this);
            return Guid.Empty;
        }
    }

    //
    // Сводка:
    //     Исходный модельный объект, по которому создана запись. Если элемент не собран
    //     по текущему документу, то возвращается null.
    public unsafe ModelObject SourceObject
    {
        get
        {
            ObjectRowElement* ptr = (ObjectRowElement*)global::_003CModule_003E.__RTDynamicCast(base.repGet, 0, System.Runtime.CompilerServices.Unsafe.AsPointer(ref global::_003CModule_003E._003F_003F_R0_003FAVRowElement_0040ps_0040_0040_00408), System.Runtime.CompilerServices.Unsafe.AsPointer(ref global::_003CModule_003E._003F_003F_R0_003FAVObjectRowElement_0040ps_0040_0040_00408), 0);
            if (ptr != null)
            {
                Document document = base.Document;
                System.Runtime.CompilerServices.Unsafe.SkipInit(out TFObjectID tFObjectID);
                ObjectId id = new ObjectId(*global::_003CModule_003E.ps_002EObjectRowElement_002EGetSourceObjectID(ptr, &tFObjectID));
                ModelObject objectById = document.GetObjectById(id);
                GC.KeepAlive(this);
                return objectById;
            }

            GC.KeepAlive(this);
            return null;
        }
    }

    //
    // Сводка:
    //     Фрагмент 3D, из которого поднят элемент. Если элемент не поднят из фрагмента
    //     или поднят из фрагмента 2D, то возвращается null.
    public unsafe ModelObject SourceFragment3DFirstLevel
    {
        get
        {
            //IL_0032: Expected I, but got I8
            AssemblyRowElement* ptr = (AssemblyRowElement*)global::_003CModule_003E.__RTDynamicCast(base.repGet, 0, System.Runtime.CompilerServices.Unsafe.AsPointer(ref global::_003CModule_003E._003F_003F_R0_003FAVRowElement_0040ps_0040_0040_00408), System.Runtime.CompilerServices.Unsafe.AsPointer(ref global::_003CModule_003E._003F_003F_R0_003FAVAssemblyRowElement_0040ps_0040_0040_00408), 0);
            if (ptr != null)
            {
                FRAGMENT* ptr2 = global::_003CModule_003E.ps_002EAssemblyRowElement_002EGetTopFragment(ptr);
                if (ptr2 != null)
                {
                    CTFObject* ptr3 = global::_003CModule_003E.FRAGMENT_002EGet3DFragment((FRAGMENT*)((ulong)(nint)ptr2 + 224uL));
                    if (ptr3 != null)
                    {
                        GC.KeepAlive(this);
                        IntPtr handle = new IntPtr(ptr3);
                        return ModelObject.FromHandle(handle);
                    }
                }
            }

            GC.KeepAlive(this);
            return null;
        }
    }

    //
    // Сводка:
    //     Фрагмент, из которого поднят элемент. Если элемент не поднят из фрагмента, то
    //     возвращается null.
    public unsafe Fragment SourceFragmentFirstLevel
    {
        get
        {
            AssemblyRowElement* ptr = (AssemblyRowElement*)global::_003CModule_003E.__RTDynamicCast(base.repGet, 0, System.Runtime.CompilerServices.Unsafe.AsPointer(ref global::_003CModule_003E._003F_003F_R0_003FAVRowElement_0040ps_0040_0040_00408), System.Runtime.CompilerServices.Unsafe.AsPointer(ref global::_003CModule_003E._003F_003F_R0_003FAVAssemblyRowElement_0040ps_0040_0040_00408), 0);
            if (ptr != null)
            {
                FRAGMENT* ptr2 = global::_003CModule_003E.ps_002EAssemblyRowElement_002EGetTopFragment(ptr);
                if (ptr2 != null)
                {
                    GC.KeepAlive(this);
                    IntPtr handle = new IntPtr(ptr2);
                    return ModelObject.FromHandle(handle) as Fragment;
                }
            }

            GC.KeepAlive(this);
            return null;
        }
        internal set
        {
            //IL_004a: Expected I, but got I8
            if (value != null && value.Document != base.Document)
            {
                throw new ArgumentOutOfRangeException("Can't set fragment from different document");
            }

            if (global::_003CModule_003E.__RTDynamicCast(base.repGet, 0, System.Runtime.CompilerServices.Unsafe.AsPointer(ref global::_003CModule_003E._003F_003F_R0_003FAVRowElement_0040ps_0040_0040_00408), System.Runtime.CompilerServices.Unsafe.AsPointer(ref global::_003CModule_003E._003F_003F_R0_003FAVAssemblyRowElement_0040ps_0040_0040_00408), 0) != null)
            {
                CTFObject* intPtr = base.repSet;
                FRAGMENT* ptr = ((value == null) ? null : value.repGet);
                global::_003CModule_003E.ps_002EAssemblyRowElement_002ESetTopFragment((AssemblyRowElement*)intPtr, ptr);
            }

            GC.KeepAlive(this);
        }
    }

    //
    // Сводка:
    //     Путь к фрагменту, из которого поднят элемент. Если элемент не поднят из фрагмента,
    //     то возвращается null.
    public unsafe string SourceFragmentPath
    {
        get
        {
            //IL_007a: Expected I, but got I8
            //IL_008b: Expected I, but got I8
            //IL_0129: Expected I, but got I8
            //IL_00d3: Expected I, but got I8
            //IL_00ef: Expected I, but got I8
            //IL_01ed: Expected I, but got I8
            //IL_024f: Expected I, but got I8
            AssemblyRowElement* ptr = (AssemblyRowElement*)global::_003CModule_003E.__RTDynamicCast(base.repGet, 0, System.Runtime.CompilerServices.Unsafe.AsPointer(ref global::_003CModule_003E._003F_003F_R0_003FAVRowElement_0040ps_0040_0040_00408), System.Runtime.CompilerServices.Unsafe.AsPointer(ref global::_003CModule_003E._003F_003F_R0_003FAVAssemblyRowElement_0040ps_0040_0040_00408), 0);
            System.Runtime.CompilerServices.Unsafe.SkipInit(out EmbeddedPath embeddedPath);
            string result;
            if (ptr != null)
            {
                CTfw32Doc* ptr2 = global::_003CModule_003E.CTFObject_002EGetDocumentPtr((CTFObject*)ptr);
                global::_003CModule_003E.std_002Evector_003Cunsigned_0020int_002Cstd_003A_003Aallocator_003Cunsigned_0020int_003E_0020_003E_002E_007Bctor_007D((vector_003Cunsigned_0020int_002Cstd_003A_003Aallocator_003Cunsigned_0020int_003E_0020_003E*)(&embeddedPath));
                try
                {
                    global::_003CModule_003E.SharedPathName_002E_007Bctor_007D((SharedPathName*)System.Runtime.CompilerServices.Unsafe.AsPointer(ref System.Runtime.CompilerServices.Unsafe.AddByteOffset(ref embeddedPath, 24)));
                }
                catch
                {
                    //try-fault
                    global::_003CModule_003E.___CxxCallUnwindDtor((delegate*<void*, void>)(delegate*<vector_003Cunsigned_0020int_002Cstd_003A_003Aallocator_003Cunsigned_0020int_003E_0020_003E*, void>)(&global::_003CModule_003E.std_002Evector_003Cunsigned_0020int_002Cstd_003A_003Aallocator_003Cunsigned_0020int_003E_0020_003E_002E_007Bdtor_007D), &embeddedPath);
                    throw;
                }

                System.Runtime.CompilerServices.Unsafe.SkipInit(out SharedPathName sharedPathName);
                System.Runtime.CompilerServices.Unsafe.SkipInit(out TFDocRegenContext tFDocRegenContext);
                System.Runtime.CompilerServices.Unsafe.SkipInit(out AssemblyNodeAutoPtr assemblyNodeAutoPtr);
                try
                {
                    global::_003CModule_003E.CTfw32Doc_002EGetEmbeddedPath(ptr2, &embeddedPath);
                    global::_003CModule_003E.SharedPathName_002E_007Bctor_007D(&sharedPathName, (SharedPathName*)System.Runtime.CompilerServices.Unsafe.AsPointer(ref System.Runtime.CompilerServices.Unsafe.AddByteOffset(ref embeddedPath, 24)));
                    try
                    {
                        global::_003CModule_003E.TFDocRegenContext_002E_007Bctor_007D(&tFDocRegenContext, ptr2, true, false, true);
                        try
                        {
                            global::_003CModule_003E.AssemblyNodeAutoPtr_002E_007Bctor_007D(&assemblyNodeAutoPtr, (AssemblyInfo.Node*)null);
                            try
                            {
                                FragRowElemID* ptr3 = global::_003CModule_003E.ps_002EAssemblyRowElement_002EGetFragRowElemID(ptr);
                                FragRowElemID* ptr4 = (FragRowElemID*)((ulong)(nint)ptr3 + 32uL);
                                uint num = (uint)global::_003CModule_003E.std_002Evector_003Cps_003A_003AFragID_002Cstd_003A_003Aallocator_003Cps_003A_003AFragID_003E_0020_003E_002Esize((vector_003Cps_003A_003AFragID_002Cstd_003A_003Aallocator_003Cps_003A_003AFragID_003E_0020_003E*)ptr4);
                                uint num2 = 0u;
                                if (0 < num)
                                {
                                    System.Runtime.CompilerServices.Unsafe.SkipInit(out TFObjectID tFObjectID);
                                    System.Runtime.CompilerServices.Unsafe.SkipInit(out AssemblyNodeAutoPtr assemblyNodeAutoPtr2);
                                    System.Runtime.CompilerServices.Unsafe.SkipInit(out AssemblyNodeAutoPtr assemblyNodeAutoPtr3);
                                    System.Runtime.CompilerServices.Unsafe.SkipInit(out TFObjectID tFObjectID2);
                                    System.Runtime.CompilerServices.Unsafe.SkipInit(out TFObjectID tFObjectID3);
                                    System.Runtime.CompilerServices.Unsafe.SkipInit(out SharedPathName sharedPathName2);
                                    while (true)
                                    {
                                        if (num2 == 0)
                                        {
                                            CTFCom_Is3D* ptr5 = global::_003CModule_003E.CTFCom_Is3D_002EInstance();
                                            long num3 = *(long*)(*(long*)ptr5 + 3000);
                                            FRAGMENT* ptr6 = ((delegate* unmanaged[Cdecl, Cdecl]<IntPtr, CTfw32Doc*, TFObjectID, FRAGMENT*>)num3)((nint)ptr5, ptr2, *global::_003CModule_003E.ps_002EFragRowElemID_002EGetTopFragID(ptr3, &tFObjectID));
                                            if (ptr6 != null)
                                            {
                                                AssemblyNodeAutoPtr* ptr7 = global::_003CModule_003E.FileLinkRef_002EassemblyNode(global::_003CModule_003E.FRAGMENT_002EFileLink((FRAGMENT*)((ulong)(nint)global::_003CModule_003E.FRAGMENT_002EGetOriginalSource(ptr6) + 224uL)), &assemblyNodeAutoPtr2, &tFDocRegenContext);
                                                try
                                                {
                                                    global::_003CModule_003E.AssemblyNodeAutoPtr_002E_003D(&assemblyNodeAutoPtr, ptr7);
                                                }
                                                catch
                                                {
                                                    //try-fault
                                                    global::_003CModule_003E.___CxxCallUnwindDtor((delegate*<void*, void>)(delegate*<AssemblyNodeAutoPtr*, void>)(&global::_003CModule_003E.AssemblyNodeAutoPtr_002E_007Bdtor_007D), &assemblyNodeAutoPtr2);
                                                    throw;
                                                }

                                                global::_003CModule_003E.AssemblyNodeAutoPtr_002E_007Bdtor_007D(&assemblyNodeAutoPtr2);
                                            }
                                        }
                                        else
                                        {
                                            global::_003CModule_003E.AssemblyNodeAutoPtr_002E_007Bctor_007D(&assemblyNodeAutoPtr3, (AssemblyInfo.Node*)null);
                                            try
                                            {
                                                vector_003CAssemblyNodeAutoPtr_002Cstd_003A_003Aallocator_003CAssemblyNodeAutoPtr_003E_0020_003E* ptr8 = global::_003CModule_003E.AssemblyInfo_002ENode_002EgetChildren(global::_003CModule_003E.AssemblyNodeAutoPtr_002E_002D_003E(&assemblyNodeAutoPtr));
                                                AssemblyNodeAutoPtr* ptr9 = global::_003CModule_003E.std_002Evector_003CAssemblyNodeAutoPtr_002Cstd_003A_003Aallocator_003CAssemblyNodeAutoPtr_003E_0020_003E_002E_Unchecked_begin(ptr8);
                                                AssemblyNodeAutoPtr* ptr10 = global::_003CModule_003E.std_002Evector_003CAssemblyNodeAutoPtr_002Cstd_003A_003Aallocator_003CAssemblyNodeAutoPtr_003E_0020_003E_002E_Unchecked_end(ptr8);
                                                if (ptr9 != ptr10)
                                                {
                                                    do
                                                    {
                                                        int num5;
                                                        bool flag;
                                                        if (!global::_003CModule_003E.AssemblyNodeAutoPtr_002Enull(ptr9))
                                                        {
                                                            ulong num4 = num2;
                                                            if (!global::_003CModule_003E.ps_002EFragID_002EIsOldID(global::_003CModule_003E.std_002Evector_003Cps_003A_003AFragID_002Cstd_003A_003Aallocator_003Cps_003A_003AFragID_003E_0020_003E_002E_005B_005D((vector_003Cps_003A_003AFragID_002Cstd_003A_003Aallocator_003Cps_003A_003AFragID_003E_0020_003E*)ptr4, num4)))
                                                            {
                                                                FragID* ptr11 = global::_003CModule_003E.std_002Evector_003Cps_003A_003AFragID_002Cstd_003A_003Aallocator_003Cps_003A_003AFragID_003E_0020_003E_002E_005B_005D((vector_003Cps_003A_003AFragID_002Cstd_003A_003Aallocator_003Cps_003A_003AFragID_003E_0020_003E*)ptr4, num4);
                                                                if (global::_003CModule_003E._003D_003D(global::_003CModule_003E.AssemblyInfo_002ENode_002EgetFragGuid(global::_003CModule_003E.AssemblyNodeAutoPtr_002E_002D_003E(ptr9)), (UniqueIdentifier*)ptr11))
                                                                {
                                                                    AssemblyInfo.Node* intPtr = global::_003CModule_003E.AssemblyNodeAutoPtr_002E_002D_003E(ptr9);
                                                                    FragRowElemID* ptr12 = ptr4;
                                                                    if (global::_003CModule_003E.TFObjectID_002EtoShortForm(global::_003CModule_003E.AssemblyInfo_002ENode_002EgetID(intPtr, &tFObjectID2)) == (uint)(*(int*)((ulong)(nint)global::_003CModule_003E.std_002Evector_003Cps_003A_003AFragID_002Cstd_003A_003Aallocator_003Cps_003A_003AFragID_003E_0020_003E_002E_005B_005D((vector_003Cps_003A_003AFragID_002Cstd_003A_003Aallocator_003Cps_003A_003AFragID_003E_0020_003E*)ptr12, num4) + 16uL)))
                                                                    {
                                                                        num5 = 1;
                                                                        goto IL_01b6;
                                                                    }
                                                                }

                                                                num5 = 0;
                                                                goto IL_01b6;
                                                            }

                                                            AssemblyInfo.Node* intPtr2 = global::_003CModule_003E.AssemblyNodeAutoPtr_002E_002D_003E(ptr9);
                                                            FragRowElemID* ptr13 = ptr4;
                                                            flag = global::_003CModule_003E.TFObjectID_002EtoShortForm(global::_003CModule_003E.AssemblyInfo_002ENode_002EgetID(intPtr2, &tFObjectID3)) == (uint)(*(int*)((ulong)(nint)global::_003CModule_003E.std_002Evector_003Cps_003A_003AFragID_002Cstd_003A_003Aallocator_003Cps_003A_003AFragID_003E_0020_003E_002E_005B_005D((vector_003Cps_003A_003AFragID_002Cstd_003A_003Aallocator_003Cps_003A_003AFragID_003E_0020_003E*)ptr13, num4) + 16uL));
                                                            goto IL_01e4;
                                                        }

                                                        goto IL_01e8;
                                                    IL_01e8:
                                                        ptr9 = (AssemblyNodeAutoPtr*)((ulong)(nint)ptr9 + 8uL);
                                                        continue;
                                                    IL_01b6:
                                                        flag = (byte)num5 != 0;
                                                        goto IL_01e4;
                                                    IL_01e4:
                                                        if (flag)
                                                        {
                                                            global::_003CModule_003E.AssemblyNodeAutoPtr_002E_003D(&assemblyNodeAutoPtr3, ptr9);
                                                            break;
                                                        }

                                                        goto IL_01e8;
                                                    }
                                                    while (ptr9 != ptr10);
                                                }

                                                global::_003CModule_003E.AssemblyNodeAutoPtr_002E_003D(&assemblyNodeAutoPtr, &assemblyNodeAutoPtr3);
                                            }
                                            catch
                                            {
                                                //try-fault
                                                global::_003CModule_003E.___CxxCallUnwindDtor((delegate*<void*, void>)(delegate*<AssemblyNodeAutoPtr*, void>)(&global::_003CModule_003E.AssemblyNodeAutoPtr_002E_007Bdtor_007D), &assemblyNodeAutoPtr3);
                                                throw;
                                            }

                                            global::_003CModule_003E.AssemblyNodeAutoPtr_002E_007Bdtor_007D(&assemblyNodeAutoPtr3);
                                        }

                                        if (!global::_003CModule_003E.AssemblyNodeAutoPtr_002Enull(&assemblyNodeAutoPtr))
                                        {
                                            StructInfo* ptr14 = global::_003CModule_003E.AssemblyInfo_002EgetInfo_003Cclass_0020AssemblyInfo_003A_003AStructInfo_003E(global::_003CModule_003E.AssemblyNodeAutoPtr_002Eptr(&assemblyNodeAutoPtr));
                                            if (ptr14 != null)
                                            {
                                                SharedPathName* ptr15 = global::_003CModule_003E.ServiceFileLinkRef_002EpathName(global::_003CModule_003E.AssemblyInfo_002EStructInfo_002EgetLink(ptr14), &sharedPathName2, &sharedPathName, (FragmentDataType)2, null);
                                                try
                                                {
                                                    global::_003CModule_003E.SharedPathName_002E_003D(&sharedPathName, ptr15);
                                                }
                                                catch
                                                {
                                                    //try-fault
                                                    global::_003CModule_003E.___CxxCallUnwindDtor((delegate*<void*, void>)(delegate*<SharedPathName*, void>)(&global::_003CModule_003E.SharedPathName_002E_007Bdtor_007D), &sharedPathName2);
                                                    throw;
                                                }

                                                global::_003CModule_003E.SharedPathName_002E_007Bdtor_007D(&sharedPathName2);
                                                num2++;
                                                if (num2 >= num)
                                                {
                                                    break;
                                                }

                                                continue;
                                            }

                                            GC.KeepAlive(this);
                                            result = null;
                                        }
                                        else
                                        {
                                            GC.KeepAlive(this);
                                            result = null;
                                        }

                                        goto IL_02a7;
                                    }
                                }
                            }
                            catch
                            {
                                //try-fault
                                global::_003CModule_003E.___CxxCallUnwindDtor((delegate*<void*, void>)(delegate*<AssemblyNodeAutoPtr*, void>)(&global::_003CModule_003E.AssemblyNodeAutoPtr_002E_007Bdtor_007D), &assemblyNodeAutoPtr);
                                throw;
                            }

                            goto end_IL_0071;
                        IL_02a7:
                            global::_003CModule_003E.AssemblyNodeAutoPtr_002E_007Bdtor_007D(&assemblyNodeAutoPtr);
                            goto IL_02be;
                        end_IL_0071:;
                        }
                        catch
                        {
                            //try-fault
                            global::_003CModule_003E.___CxxCallUnwindDtor((delegate*<void*, void>)(delegate*<TFDocRegenContext*, void>)(&global::_003CModule_003E.TFDocRegenContext_002E_007Bdtor_007D), &tFDocRegenContext);
                            throw;
                        }

                        goto end_IL_0064;
                    IL_02be:
                        global::_003CModule_003E.TFDocRegenContext_002E_007Bdtor_007D(&tFDocRegenContext);
                        goto IL_02d5;
                    end_IL_0064:;
                    }
                    catch
                    {
                        //try-fault
                        global::_003CModule_003E.___CxxCallUnwindDtor((delegate*<void*, void>)(delegate*<SharedPathName*, void>)(&global::_003CModule_003E.SharedPathName_002E_007Bdtor_007D), &sharedPathName);
                        throw;
                    }

                    goto end_IL_004d;
                IL_02d5:
                    global::_003CModule_003E.SharedPathName_002E_007Bdtor_007D(&sharedPathName);
                    goto IL_02ed;
                end_IL_004d:;
                }
                catch
                {
                    //try-fault
                    global::_003CModule_003E.___CxxCallUnwindDtor((delegate*<void*, void>)(delegate*<EmbeddedPath*, void>)(&global::_003CModule_003E.EmbeddedPath_002E_007Bdtor_007D), &embeddedPath);
                    throw;
                }

                string result2;
                try
                {
                    try
                    {
                        try
                        {
                            try
                            {
                                result2 = new string(global::_003CModule_003E.SharedPathName_002EasString(&sharedPathName));
                            }
                            catch
                            {
                                //try-fault
                                global::_003CModule_003E.___CxxCallUnwindDtor((delegate*<void*, void>)(delegate*<AssemblyNodeAutoPtr*, void>)(&global::_003CModule_003E.AssemblyNodeAutoPtr_002E_007Bdtor_007D), &assemblyNodeAutoPtr);
                                throw;
                            }

                            global::_003CModule_003E.AssemblyNodeAutoPtr_002E_007Bdtor_007D(&assemblyNodeAutoPtr);
                        }
                        catch
                        {
                            //try-fault
                            global::_003CModule_003E.___CxxCallUnwindDtor((delegate*<void*, void>)(delegate*<TFDocRegenContext*, void>)(&global::_003CModule_003E.TFDocRegenContext_002E_007Bdtor_007D), &tFDocRegenContext);
                            throw;
                        }

                        global::_003CModule_003E.TFDocRegenContext_002E_007Bdtor_007D(&tFDocRegenContext);
                    }
                    catch
                    {
                        //try-fault
                        global::_003CModule_003E.___CxxCallUnwindDtor((delegate*<void*, void>)(delegate*<SharedPathName*, void>)(&global::_003CModule_003E.SharedPathName_002E_007Bdtor_007D), &sharedPathName);
                        throw;
                    }

                    global::_003CModule_003E.SharedPathName_002E_007Bdtor_007D(&sharedPathName);
                }
                catch
                {
                    //try-fault
                    global::_003CModule_003E.___CxxCallUnwindDtor((delegate*<void*, void>)(delegate*<EmbeddedPath*, void>)(&global::_003CModule_003E.EmbeddedPath_002E_007Bdtor_007D), &embeddedPath);
                    throw;
                }

                try
                {
                    global::_003CModule_003E.SharedPathName_002E_007Bdtor_007D((SharedPathName*)System.Runtime.CompilerServices.Unsafe.AsPointer(ref System.Runtime.CompilerServices.Unsafe.AddByteOffset(ref embeddedPath, 24)));
                }
                catch
                {
                    //try-fault
                    global::_003CModule_003E.___CxxCallUnwindDtor((delegate*<void*, void>)(delegate*<vector_003Cunsigned_0020int_002Cstd_003A_003Aallocator_003Cunsigned_0020int_003E_0020_003E*, void>)(&global::_003CModule_003E.std_002Evector_003Cunsigned_0020int_002Cstd_003A_003Aallocator_003Cunsigned_0020int_003E_0020_003E_002E_007Bdtor_007D), &embeddedPath);
                    throw;
                }

                global::_003CModule_003E.std_002Evector_003Cunsigned_0020int_002Cstd_003A_003Aallocator_003Cunsigned_0020int_003E_0020_003E_002E_007Bdtor_007D((vector_003Cunsigned_0020int_002Cstd_003A_003Aallocator_003Cunsigned_0020int_003E_0020_003E*)(&embeddedPath));
                GC.KeepAlive(this);
                return result2;
            }

            GC.KeepAlive(this);
            return null;
        IL_02ed:
            try
            {
                global::_003CModule_003E.SharedPathName_002E_007Bdtor_007D((SharedPathName*)System.Runtime.CompilerServices.Unsafe.AsPointer(ref System.Runtime.CompilerServices.Unsafe.AddByteOffset(ref embeddedPath, 24)));
            }
            catch
            {
                //try-fault
                global::_003CModule_003E.___CxxCallUnwindDtor((delegate*<void*, void>)(delegate*<vector_003Cunsigned_0020int_002Cstd_003A_003Aallocator_003Cunsigned_0020int_003E_0020_003E*, void>)(&global::_003CModule_003E.std_002Evector_003Cunsigned_0020int_002Cstd_003A_003Aallocator_003Cunsigned_0020int_003E_0020_003E_002E_007Bdtor_007D), &embeddedPath);
                throw;
            }

            global::_003CModule_003E.std_002Evector_003Cunsigned_0020int_002Cstd_003A_003Aallocator_003Cunsigned_0020int_003E_0020_003E_002E_007Bdtor_007D((vector_003Cunsigned_0020int_002Cstd_003A_003Aallocator_003Cunsigned_0020int_003E_0020_003E*)(&embeddedPath));
            return result;
        }
    }

    //
    // Сводка:
    //     Родительский элемент
    public unsafe RowElement ParentRowElement
    {
        get
        {
            ps.RowElement* ptr = global::_003CModule_003E.ps_002ERowElement_002EGetParentRowElement((ps.RowElement*)base.repGet);
            if (ptr != null)
            {
                GC.KeepAlive(this);
                return ModelObject.FromHandle((IntPtr)ptr) as RowElement;
            }

            GC.KeepAlive(this);
            return null;
        }
        set
        {
            if (value != null && value.ProductStructure != ProductStructure)
            {
                throw new ArgumentException("invalid parentRowElem");
            }

            _ = ProductStructure.repSet;
            CTFObject* ptr = base.repSet;
            System.Runtime.CompilerServices.Unsafe.SkipInit(out StructureManager structureManager);
            StructureManager* intPtr = global::_003CModule_003E.ps_002EStructureManager_002E_007Bctor_007D(&structureManager, global::_003CModule_003E.CTFObject_002EGetDocumentPtr(ptr), true);
            int num = ((value != null) ? global::_003CModule_003E.CTFObject_002EGetIndexInModelContainer(value.repGet) : (-1));
            int num2 = global::_003CModule_003E.ps_002ERowElement_002EGetIndex((ps.RowElement*)ptr);
            global::_003CModule_003E.ps_002EStructureManager_002ESetRowElemParent(intPtr, num2, num, true);
            GC.KeepAlive(this);
        }
    }

    //
    // Сводка:
    //     Структура изделия, которой принадлежит элемент
    public unsafe ProductStructure ProductStructure
    {
        get
        {
            CTFObject* ptr = base.repGet;
            if (ptr != null)
            {
                ps.ProductStructure* ptr2 = global::_003CModule_003E.ps_002ERowElement_002EGetProductStruct((ps.RowElement*)ptr);
                if (ptr2 != null)
                {
                    GC.KeepAlive(this);
                    return ModelObject.FromHandle((IntPtr)ptr2) as ProductStructure;
                }
            }

            GC.KeepAlive(this);
            return null;
        }
    }

    //
    // Сводка:
    //     Уникальный идентификатор элемента
    public unsafe Guid UID
    {
        get
        {
            CTFObject* ptr = base.repGet;
            UniqueIdentifier uniqueIdentifier = *global::_003CModule_003E.ps_002ERowElement_002EGetUID((ps.RowElement*)ptr);
            Guid result = global::_003CModule_003E.TFlex_002EGUID2Guid(global::_003CModule_003E.UniqueIdentifier_002E_002EAEBU_GUID_0040_0040(&uniqueIdentifier));
            GC.KeepAlive(this);
            return result;
        }
    }

    internal unsafe RowElement(ps.RowElement* pNewElement)
        : base(new IntPtr(pNewElement), addToContainer: true)
    {
    }

    internal RowElement(IntPtr Handle)
        : base(Handle, addToContainer: false)
    {
    }

    //
    // Сводка:
    //     Список идентификаторов фрагментов, из которых поднят элемент. Если элемент не
    //     поднят из фрагмента, то возвращается пустой перечислитель.
    //
    // Параметры:
    //   fragment3d:
    //     Возвращать идентификаторы для 3D фрагментов
    public unsafe IEnumerable<ObjectId> GetSourceFragmentIdChain([MarshalAs(UnmanagedType.U1)] bool fragment3d)
    {
        //IL_001e: Expected I, but got I8
        //IL_0020: Expected I8, but got I
        //IL_0029: Expected I, but got I8
        //IL_009e: Expected I, but got I8
        CTFObject* ptr = base.repGet;
        long num = *(long*)(*(long*)ptr + 2016);
        System.Runtime.CompilerServices.Unsafe.SkipInit(out optional_003CTFObjectIDVector_003E optional_003CTFObjectIDVector_003E);
        long num2 = (nint)((delegate* unmanaged[Cdecl, Cdecl]<IntPtr, optional_003CTFObjectIDVector_003E*, byte, optional_003CTFObjectIDVector_003E*>)num)((nint)ptr, &optional_003CTFObjectIDVector_003E, 1);
        System.Runtime.CompilerServices.Unsafe.SkipInit(out optional_003CTFObjectIDVector_003E optional_003CTFObjectIDVector_003E2);
        try
        {
            global::_003CModule_003E.std_002E_Non_trivial_move_003Cstd_003A_003A_Optional_construct_base_003CTFObjectIDVector_003E_002CTFObjectIDVector_003E_002E_007Bctor_007D((_Non_trivial_move_003Cstd_003A_003A_Optional_construct_base_003CTFObjectIDVector_003E_002CTFObjectIDVector_003E*)(&optional_003CTFObjectIDVector_003E2), (_Non_trivial_move_003Cstd_003A_003A_Optional_construct_base_003CTFObjectIDVector_003E_002CTFObjectIDVector_003E*)num2);
        }
        catch
        {
            //try-fault
            global::_003CModule_003E.___CxxCallUnwindDtor((delegate*<void*, void>)(delegate*<optional_003CTFObjectIDVector_003E*, void>)(&global::_003CModule_003E.std_002Eoptional_003CTFObjectIDVector_003E_002E_007Bdtor_007D), &optional_003CTFObjectIDVector_003E);
            throw;
        }

        IEnumerable<ObjectId> result;
        IEnumerable<ObjectId> result2;
        try
        {
            global::_003CModule_003E.std_002E_Optional_destruct_base_003CTFObjectIDVector_002C0_003E_002E_007Bdtor_007D((_Optional_destruct_base_003CTFObjectIDVector_002C0_003E*)(&optional_003CTFObjectIDVector_003E));
            if (!global::_003CModule_003E.std_002Eoptional_003CTFObjectIDVector_003E_002Ehas_value(&optional_003CTFObjectIDVector_003E2))
            {
                GC.KeepAlive(this);
                result = Enumerable.Empty<ObjectId>();
                goto IL_0119;
            }

            if (fragment3d)
            {
                CTFCom_Is3D* ptr2 = global::_003CModule_003E.CTFCom_Is3D_002EInstance();
                long num3 = *(long*)(*(long*)ptr2 + 3000);
                CTfw32Doc* ptr3 = global::_003CModule_003E.CTFObject_002EGetDocumentPtr(base.repGet);
                FRAGMENT* ptr4 = ((delegate* unmanaged[Cdecl, Cdecl]<IntPtr, CTfw32Doc*, TFObjectID, FRAGMENT*>)num3)((nint)ptr2, ptr3, *global::_003CModule_003E.std_002Evector_003CTFObjectID_002Cstd_003A_003Aallocator_003CTFObjectID_003E_0020_003E_002Efront((vector_003CTFObjectID_002Cstd_003A_003Aallocator_003CTFObjectID_003E_0020_003E*)global::_003CModule_003E.std_002Eoptional_003CTFObjectIDVector_003E_002E_002D_003E(&optional_003CTFObjectIDVector_003E2)));
                if (ptr4 == null)
                {
                    goto IL_00fc;
                }

                System.Runtime.CompilerServices.Unsafe.SkipInit(out TFObjectIDVector tFObjectIDVector);
                global::_003CModule_003E.TFObjectIDVector_002E_007Bctor_007D(&tFObjectIDVector);
                try
                {
                    global::_003CModule_003E.ps_002Econvert2Dto3DVecID(ptr4, global::_003CModule_003E.std_002Eoptional_003CTFObjectIDVector_003E_002E_002A(&optional_003CTFObjectIDVector_003E2), &tFObjectIDVector);
                    result2 = global::_003CModule_003E.TFlex_002EModel_002E_003FA0xadee7240_002EtoManaged(&tFObjectIDVector);
                }
                catch
                {
                    //try-fault
                    global::_003CModule_003E.___CxxCallUnwindDtor((delegate*<void*, void>)(delegate*<TFObjectIDVector*, void>)(&global::_003CModule_003E.TFObjectIDVector_002E_007Bdtor_007D), &tFObjectIDVector);
                    throw;
                }

                global::_003CModule_003E.TFObjectIDVector_002E_007Bdtor_007D(&tFObjectIDVector);
                goto IL_00eb;
            }
        }
        catch
        {
            //try-fault
            global::_003CModule_003E.___CxxCallUnwindDtor((delegate*<void*, void>)(delegate*<optional_003CTFObjectIDVector_003E*, void>)(&global::_003CModule_003E.std_002Eoptional_003CTFObjectIDVector_003E_002E_007Bdtor_007D), &optional_003CTFObjectIDVector_003E2);
            throw;
        }

        IEnumerable<ObjectId> result3;
        try
        {
            result3 = global::_003CModule_003E.TFlex_002EModel_002E_003FA0xadee7240_002EtoManaged(global::_003CModule_003E.std_002Eoptional_003CTFObjectIDVector_003E_002E_002A(&optional_003CTFObjectIDVector_003E2));
        }
        catch
        {
            //try-fault
            global::_003CModule_003E.___CxxCallUnwindDtor((delegate*<void*, void>)(delegate*<optional_003CTFObjectIDVector_003E*, void>)(&global::_003CModule_003E.std_002Eoptional_003CTFObjectIDVector_003E_002E_007Bdtor_007D), &optional_003CTFObjectIDVector_003E2);
            throw;
        }

        global::_003CModule_003E.std_002E_Optional_destruct_base_003CTFObjectIDVector_002C0_003E_002E_007Bdtor_007D((_Optional_destruct_base_003CTFObjectIDVector_002C0_003E*)(&optional_003CTFObjectIDVector_003E2));
        GC.KeepAlive(this);
        return result3;
    IL_0119:
        try
        {
        }
        catch
        {
            //try-fault
            global::_003CModule_003E.___CxxCallUnwindDtor((delegate*<void*, void>)(delegate*<optional_003CTFObjectIDVector_003E*, void>)(&global::_003CModule_003E.std_002Eoptional_003CTFObjectIDVector_003E_002E_007Bdtor_007D), &optional_003CTFObjectIDVector_003E2);
            throw;
        }

        global::_003CModule_003E.std_002E_Optional_destruct_base_003CTFObjectIDVector_002C0_003E_002E_007Bdtor_007D((_Optional_destruct_base_003CTFObjectIDVector_002C0_003E*)(&optional_003CTFObjectIDVector_003E2));
        return result;
    IL_00eb:
        global::_003CModule_003E.std_002E_Optional_destruct_base_003CTFObjectIDVector_002C0_003E_002E_007Bdtor_007D((_Optional_destruct_base_003CTFObjectIDVector_002C0_003E*)(&optional_003CTFObjectIDVector_003E2));
        GC.KeepAlive(this);
        return result2;
    IL_00fc:
        try
        {
            GC.KeepAlive(this);
            result = Enumerable.Empty<ObjectId>();
        }
        catch
        {
            //try-fault
            global::_003CModule_003E.___CxxCallUnwindDtor((delegate*<void*, void>)(delegate*<optional_003CTFObjectIDVector_003E*, void>)(&global::_003CModule_003E.std_002Eoptional_003CTFObjectIDVector_003E_002E_007Bdtor_007D), &optional_003CTFObjectIDVector_003E2);
            throw;
        }

        goto IL_0119;
    }

    //
    // Сводка:
    //     Получить ячейку элемента для колонки структуры изделия
    //
    // Параметры:
    //   parameterId:
    //     Идентификатор колонки структуры изделия
    public RowElementCell GetCell(Guid parameterId)
    {
        return new RowElementCell(this, parameterId);
    }

    //
    // Сводка:
    //     Получить ячейку элемента для колонки структуры изделия
    //
    // Параметры:
    //   parameter:
    //     Колонка структуры изделия
    public RowElementCell GetCell(TFlex.Model.Data.ProductStructure.ParameterDescriptor parameter)
    {
        Guid uID = parameter.UID;
        return new RowElementCell(this, uID);
    }

    //
    // Сводка:
    //     Создать область изменений. Запись не будет автоматически пересчитываться при
    //     изменениях в ячейках. Вместо этого она пересчитается когда будет вызван Dispose.
    //
    //
    // Примечания:
    //     Для использования в using.
    public IDisposable CreateChangesScope()
    {
        return new ModelObjectChangesScope(this, regenerate: true);
    }

    internal unsafe object GetCellValue(Guid parameterId)
    {
        CTFObject* ptr = base.repGet;
        DataTable* dataTable = GetDataTable(bForSet: false);
        ps.Scheme* intPtr = global::_003CModule_003E.ps_002EProductStructure_002EGetScheme(global::_003CModule_003E.ps_002ERowElement_002EGetProductStruct((ps.RowElement*)ptr));
        System.Runtime.CompilerServices.Unsafe.SkipInit(out _GUID gUID);
        IntPtr ptr2 = new IntPtr(&gUID);
        Marshal.StructureToPtr(parameterId, ptr2, fDeleteOld: false);
        _GUID gUID2 = gUID;
        System.Runtime.CompilerServices.Unsafe.SkipInit(out UniqueIdentifier uniqueIdentifier);
        global::_003CModule_003E.UniqueIdentifier_002E_007Bctor_007D(&uniqueIdentifier, &gUID2);
        ps.ParameterDescriptor* ptr3 = global::_003CModule_003E.ps_002EScheme_002EGetParameterDescriptorWithAux(intPtr, &uniqueIdentifier);
        if (ptr3 == null)
        {
            throw new ArgumentException("wrong parameterId for cell");
        }

        System.Runtime.CompilerServices.Unsafe.SkipInit(out CStringT_003Cwchar_t_002CStrTraitMFC_DLL_003Cwchar_t_002CATL_003A_003AChTraitsCRT_003Cwchar_t_003E_0020_003E_0020_003E obj);
        object result;
        switch (global::_003CModule_003E.ps_002EParameterDescriptor_002EGetValueType(ptr3))
        {
            default:
                GC.KeepAlive(this);
                return null;
            case (BomValue.EType)4:
                {
                    global::_003CModule_003E.ATL_002ECStringT_003Cwchar_t_002CStrTraitMFC_DLL_003Cwchar_t_002CATL_003A_003AChTraitsCRT_003Cwchar_t_003E_0020_003E_0020_003E_002E_007Bctor_007D(&obj);
                    try
                    {
                        if (global::_003CModule_003E.ps_002EDataTable_002EGetValue(dataTable, global::_003CModule_003E.ps_002ERowElement_002EGetUID((ps.RowElement*)ptr), &uniqueIdentifier, &obj))
                        {
                            result = global::_003CModule_003E.TFlex_002EToNetString(&obj);
                            goto IL_00b6;
                        }
                    }
                    catch
                    {
                        //try-fault
                        global::_003CModule_003E.___CxxCallUnwindDtor((delegate*<void*, void>)(delegate*<CStringT_003Cwchar_t_002CStrTraitMFC_DLL_003Cwchar_t_002CATL_003A_003AChTraitsCRT_003Cwchar_t_003E_0020_003E_0020_003E*, void>)(&global::_003CModule_003E.ATL_002ECStringT_003Cwchar_t_002CStrTraitMFC_DLL_003Cwchar_t_002CATL_003A_003AChTraitsCRT_003Cwchar_t_003E_0020_003E_0020_003E_002E_007Bdtor_007D), &obj);
                        throw;
                    }

                    object result2;
                    try
                    {
                        GC.KeepAlive(this);
                        result2 = null;
                    }
                    catch
                    {
                        //try-fault
                        global::_003CModule_003E.___CxxCallUnwindDtor((delegate*<void*, void>)(delegate*<CStringT_003Cwchar_t_002CStrTraitMFC_DLL_003Cwchar_t_002CATL_003A_003AChTraitsCRT_003Cwchar_t_003E_0020_003E_0020_003E*, void>)(&global::_003CModule_003E.ATL_002ECStringT_003Cwchar_t_002CStrTraitMFC_DLL_003Cwchar_t_002CATL_003A_003AChTraitsCRT_003Cwchar_t_003E_0020_003E_0020_003E_002E_007Bdtor_007D), &obj);
                        throw;
                    }

                    global::_003CModule_003E.ATL_002ECStringT_003Cwchar_t_002CStrTraitMFC_DLL_003Cwchar_t_002CATL_003A_003AChTraitsCRT_003Cwchar_t_003E_0020_003E_0020_003E_002E_007Bdtor_007D(&obj);
                    return result2;
                }
            case (BomValue.EType)3:
                {
                    System.Runtime.CompilerServices.Unsafe.SkipInit(out double num);
                    if (global::_003CModule_003E.ps_002EDataTable_002EGetValue(dataTable, global::_003CModule_003E.ps_002ERowElement_002EGetUID((ps.RowElement*)ptr), &uniqueIdentifier, &num))
                    {
                        GC.KeepAlive(this);
                        return num;
                    }

                    GC.KeepAlive(this);
                    return null;
                }
            case (BomValue.EType)2:
                {
                    System.Runtime.CompilerServices.Unsafe.SkipInit(out int num2);
                    if (global::_003CModule_003E.ps_002EDataTable_002EGetValue(dataTable, global::_003CModule_003E.ps_002ERowElement_002EGetUID((ps.RowElement*)ptr), &uniqueIdentifier, &num2))
                    {
                        GC.KeepAlive(this);
                        return num2;
                    }

                    GC.KeepAlive(this);
                    return null;
                }
            case (BomValue.EType)1:
                {
                    System.Runtime.CompilerServices.Unsafe.SkipInit(out bool flag);
                    if (global::_003CModule_003E.ps_002EDataTable_002EGetValue(dataTable, global::_003CModule_003E.ps_002ERowElement_002EGetUID((ps.RowElement*)ptr), &uniqueIdentifier, &flag))
                    {
                        GC.KeepAlive(this);
                        return flag;
                    }

                    GC.KeepAlive(this);
                    return null;
                }

            IL_00b6:
                global::_003CModule_003E.ATL_002ECStringT_003Cwchar_t_002CStrTraitMFC_DLL_003Cwchar_t_002CATL_003A_003AChTraitsCRT_003Cwchar_t_003E_0020_003E_0020_003E_002E_007Bdtor_007D(&obj);
                GC.KeepAlive(this);
                return result;
        }
    }

    internal unsafe void SetCellValue(Guid parameterId, object value)
    {
        uint num = 0u;
        _ = ProductStructure.repSet;
        CTFObject* ptr = base.repGet;
        System.Runtime.CompilerServices.Unsafe.SkipInit(out StructureManager structureManager);
        global::_003CModule_003E.ps_002EStructureManager_002E_007Bctor_007D(&structureManager, global::_003CModule_003E.CTFObject_002EGetDocumentPtr(ptr), true);
        ps.Scheme* intPtr = global::_003CModule_003E.ps_002EProductStructure_002EGetScheme(global::_003CModule_003E.ps_002ERowElement_002EGetProductStruct((ps.RowElement*)ptr));
        System.Runtime.CompilerServices.Unsafe.SkipInit(out _GUID gUID);
        IntPtr ptr2 = new IntPtr(&gUID);
        Marshal.StructureToPtr(parameterId, ptr2, fDeleteOld: false);
        _GUID gUID2 = gUID;
        System.Runtime.CompilerServices.Unsafe.SkipInit(out UniqueIdentifier uniqueIdentifier);
        global::_003CModule_003E.UniqueIdentifier_002E_007Bctor_007D(&uniqueIdentifier, &gUID2);
        ps.ParameterDescriptor* ptr3 = global::_003CModule_003E.ps_002EScheme_002EGetParameterDescriptorWithAux(intPtr, &uniqueIdentifier);
        if (ptr3 == null)
        {
            throw new ArgumentException("wrong parameterId for cell");
        }

        switch (global::_003CModule_003E.ps_002EParameterDescriptor_002EGetValueType(ptr3))
        {
            default:
                GC.KeepAlive(this);
                throw new NotSupportedException("wrong column type");
            case (BomValue.EType)4:
                {
                    string str = (string)value;
                    System.Runtime.CompilerServices.Unsafe.SkipInit(out CStringT_003Cwchar_t_002CStrTraitMFC_DLL_003Cwchar_t_002CATL_003A_003AChTraitsCRT_003Cwchar_t_003E_0020_003E_0020_003E obj);
                    CStringT_003Cwchar_t_002CStrTraitMFC_DLL_003Cwchar_t_002CATL_003A_003AChTraitsCRT_003Cwchar_t_003E_0020_003E_0020_003E* pThis = &obj;
                    CStringT_003Cwchar_t_002CStrTraitMFC_DLL_003Cwchar_t_002CATL_003A_003AChTraitsCRT_003Cwchar_t_003E_0020_003E_0020_003E* ptr4 = global::_003CModule_003E.TFlex_002EToCString(&obj, str);
                    int num4;
                    try
                    {
                        num4 = global::_003CModule_003E.ps_002ERowElement_002EGetIndex((ps.RowElement*)ptr);
                    }
                    catch
                    {
                        //try-fault
                        global::_003CModule_003E.___CxxCallUnwindDtor((delegate*<void*, void>)(delegate*<CStringT_003Cwchar_t_002CStrTraitMFC_DLL_003Cwchar_t_002CATL_003A_003AChTraitsCRT_003Cwchar_t_003E_0020_003E_0020_003E*, void>)(&global::_003CModule_003E.ATL_002ECStringT_003Cwchar_t_002CStrTraitMFC_DLL_003Cwchar_t_002CATL_003A_003AChTraitsCRT_003Cwchar_t_003E_0020_003E_0020_003E_002E_007Bdtor_007D), pThis);
                        throw;
                    }

                    global::_003CModule_003E.ps_002EStructureManager_002EChangeRowElemParamDepValue(&structureManager, num4, &uniqueIdentifier, ptr4, false, -1, false, -1);
                    break;
                }
            case (BomValue.EType)3:
                {
                    double num5 = (double)value;
                    global::_003CModule_003E.ps_002EStructureManager_002EChangeRowElemParamDepValue(&structureManager, global::_003CModule_003E.ps_002ERowElement_002EGetIndex((ps.RowElement*)ptr), &uniqueIdentifier, num5, false, -1, false, -1);
                    break;
                }
            case (BomValue.EType)2:
                {
                    int num6 = (int)value;
                    global::_003CModule_003E.ps_002EStructureManager_002EChangeRowElemParamDepValue(&structureManager, global::_003CModule_003E.ps_002ERowElement_002EGetIndex((ps.RowElement*)ptr), &uniqueIdentifier, num6, false, -1, false, -1);
                    break;
                }
            case (BomValue.EType)1:
                {
                    double num2 = ((!(bool)value) ? 0.0 : 1.0);
                    int num3 = global::_003CModule_003E.ps_002ERowElement_002EGetIndex((ps.RowElement*)ptr);
                    global::_003CModule_003E.ps_002EStructureManager_002EChangeRowElemParamBoolDepValue(&structureManager, num3, &uniqueIdentifier, num2, false, -1, false, -1);
                    break;
                }
        }
    }

    internal unsafe string GetCellValueAsString(Guid parameterId)
    {
        CTFObject* ptr = base.repGet;
        DataTable* dataTable = GetDataTable(bForSet: false);
        System.Runtime.CompilerServices.Unsafe.SkipInit(out CStringT_003Cwchar_t_002CStrTraitMFC_DLL_003Cwchar_t_002CATL_003A_003AChTraitsCRT_003Cwchar_t_003E_0020_003E_0020_003E obj);
        global::_003CModule_003E.ATL_002ECStringT_003Cwchar_t_002CStrTraitMFC_DLL_003Cwchar_t_002CATL_003A_003AChTraitsCRT_003Cwchar_t_003E_0020_003E_0020_003E_002E_007Bctor_007D(&obj);
        string result;
        try
        {
            System.Runtime.CompilerServices.Unsafe.SkipInit(out _GUID gUID);
            IntPtr ptr2 = new IntPtr(&gUID);
            Marshal.StructureToPtr(parameterId, ptr2, fDeleteOld: false);
            _GUID gUID2 = gUID;
            System.Runtime.CompilerServices.Unsafe.SkipInit(out UniqueIdentifier uniqueIdentifier);
            global::_003CModule_003E.UniqueIdentifier_002E_007Bctor_007D(&uniqueIdentifier, &gUID2);
            UniqueIdentifier* ptr3 = global::_003CModule_003E.ps_002ERowElement_002EGetUID((ps.RowElement*)ptr);
            global::_003CModule_003E.ps_002EDataTable_002EGetValueAsString(dataTable, ptr3, &uniqueIdentifier, &obj);
            result = new string(global::_003CModule_003E.ATL_002ECSimpleStringT_003Cwchar_t_002C1_003E_002E_002EPEB_W((CSimpleStringT_003Cwchar_t_002C1_003E*)(&obj)));
        }
        catch
        {
            //try-fault
            global::_003CModule_003E.___CxxCallUnwindDtor((delegate*<void*, void>)(delegate*<CStringT_003Cwchar_t_002CStrTraitMFC_DLL_003Cwchar_t_002CATL_003A_003AChTraitsCRT_003Cwchar_t_003E_0020_003E_0020_003E*, void>)(&global::_003CModule_003E.ATL_002ECStringT_003Cwchar_t_002CStrTraitMFC_DLL_003Cwchar_t_002CATL_003A_003AChTraitsCRT_003Cwchar_t_003E_0020_003E_0020_003E_002E_007Bdtor_007D), &obj);
            throw;
        }

        global::_003CModule_003E.ATL_002ECStringT_003Cwchar_t_002CStrTraitMFC_DLL_003Cwchar_t_002CATL_003A_003AChTraitsCRT_003Cwchar_t_003E_0020_003E_0020_003E_002E_007Bdtor_007D(&obj);
        GC.KeepAlive(this);
        return result;
    }

    internal unsafe void SetCellValueAsString(Guid parameterId, string value)
    {
        uint num = 0u;
        _ = ProductStructure.repSet;
        CTFObject* ptr = base.repGet;
        System.Runtime.CompilerServices.Unsafe.SkipInit(out StructureManager structureManager);
        global::_003CModule_003E.ps_002EStructureManager_002E_007Bctor_007D(&structureManager, global::_003CModule_003E.CTFObject_002EGetDocumentPtr(ptr), true);
        System.Runtime.CompilerServices.Unsafe.SkipInit(out CStringT_003Cwchar_t_002CStrTraitMFC_DLL_003Cwchar_t_002CATL_003A_003AChTraitsCRT_003Cwchar_t_003E_0020_003E_0020_003E obj);
        CStringT_003Cwchar_t_002CStrTraitMFC_DLL_003Cwchar_t_002CATL_003A_003AChTraitsCRT_003Cwchar_t_003E_0020_003E_0020_003E* pThis = &obj;
        CStringT_003Cwchar_t_002CStrTraitMFC_DLL_003Cwchar_t_002CATL_003A_003AChTraitsCRT_003Cwchar_t_003E_0020_003E_0020_003E* ptr2 = global::_003CModule_003E.TFlex_002EToCString(&obj, value);
        System.Runtime.CompilerServices.Unsafe.SkipInit(out UniqueIdentifier uniqueIdentifier);
        int num2;
        try
        {
            System.Runtime.CompilerServices.Unsafe.SkipInit(out _GUID gUID);
            IntPtr ptr3 = new IntPtr(&gUID);
            Marshal.StructureToPtr(parameterId, ptr3, fDeleteOld: false);
            _GUID gUID2 = gUID;
            global::_003CModule_003E.UniqueIdentifier_002E_007Bctor_007D(&uniqueIdentifier, &gUID2);
            num2 = global::_003CModule_003E.ps_002ERowElement_002EGetIndex((ps.RowElement*)ptr);
        }
        catch
        {
            //try-fault
            global::_003CModule_003E.___CxxCallUnwindDtor((delegate*<void*, void>)(delegate*<CStringT_003Cwchar_t_002CStrTraitMFC_DLL_003Cwchar_t_002CATL_003A_003AChTraitsCRT_003Cwchar_t_003E_0020_003E_0020_003E*, void>)(&global::_003CModule_003E.ATL_002ECStringT_003Cwchar_t_002CStrTraitMFC_DLL_003Cwchar_t_002CATL_003A_003AChTraitsCRT_003Cwchar_t_003E_0020_003E_0020_003E_002E_007Bdtor_007D), pThis);
            throw;
        }

        global::_003CModule_003E.ps_002EStructureManager_002EChangeRowElemParamDepValueAsString(&structureManager, num2, &uniqueIdentifier, ptr2, false, -1, false, -1);
        GC.KeepAlive(this);
    }

    internal unsafe string GetCellText(Guid parameterId)
    {
        CTFObject* ptr = base.repGet;
        DataTable* dataTable = GetDataTable(bForSet: false);
        System.Runtime.CompilerServices.Unsafe.SkipInit(out CCompoundString cCompoundString);
        global::_003CModule_003E.CCompoundString_002E_007Bctor_007D(&cCompoundString);
        string result;
        try
        {
            System.Runtime.CompilerServices.Unsafe.SkipInit(out _GUID gUID);
            IntPtr ptr2 = new IntPtr(&gUID);
            Marshal.StructureToPtr(parameterId, ptr2, fDeleteOld: false);
            _GUID gUID2 = gUID;
            System.Runtime.CompilerServices.Unsafe.SkipInit(out UniqueIdentifier uniqueIdentifier);
            global::_003CModule_003E.UniqueIdentifier_002E_007Bctor_007D(&uniqueIdentifier, &gUID2);
            UniqueIdentifier* ptr3 = global::_003CModule_003E.ps_002ERowElement_002EGetUID((ps.RowElement*)ptr);
            if (global::_003CModule_003E.ps_002EDataTable_002EGetCCompString(dataTable, ptr3, &uniqueIdentifier, &cCompoundString))
            {
                result = global::_003CModule_003E.TFlex_002EToNetString(global::_003CModule_003E.CCompoundString_002EGetString(&cCompoundString));
                goto IL_006d;
            }
        }
        catch
        {
            //try-fault
            global::_003CModule_003E.___CxxCallUnwindDtor((delegate*<void*, void>)(delegate*<CCompoundString*, void>)(&global::_003CModule_003E.CCompoundString_002E_007Bdtor_007D), &cCompoundString);
            throw;
        }

        string empty;
        try
        {
            GC.KeepAlive(this);
            empty = string.Empty;
        }
        catch
        {
            //try-fault
            global::_003CModule_003E.___CxxCallUnwindDtor((delegate*<void*, void>)(delegate*<CCompoundString*, void>)(&global::_003CModule_003E.CCompoundString_002E_007Bdtor_007D), &cCompoundString);
            throw;
        }

        global::_003CModule_003E.CCompoundString_002E_007Bdtor_007D(&cCompoundString);
        return empty;
    IL_006d:
        global::_003CModule_003E.CCompoundString_002E_007Bdtor_007D(&cCompoundString);
        GC.KeepAlive(this);
        return result;
    }

    internal unsafe void SetCellText(Guid parameterId, string text)
    {
        uint num = 0u;
        if (text == null)
        {
            text = string.Empty;
        }

        _ = ProductStructure.repSet;
        CTFObject* ptr = base.repGet;
        System.Runtime.CompilerServices.Unsafe.SkipInit(out StructureManager structureManager);
        StructureManager* ptr2 = global::_003CModule_003E.ps_002EStructureManager_002E_007Bctor_007D(&structureManager, global::_003CModule_003E.CTFObject_002EGetDocumentPtr(ptr), true);
        System.Runtime.CompilerServices.Unsafe.SkipInit(out CStringT_003Cwchar_t_002CStrTraitMFC_DLL_003Cwchar_t_002CATL_003A_003AChTraitsCRT_003Cwchar_t_003E_0020_003E_0020_003E obj);
        CStringT_003Cwchar_t_002CStrTraitMFC_DLL_003Cwchar_t_002CATL_003A_003AChTraitsCRT_003Cwchar_t_003E_0020_003E_0020_003E* pThis = &obj;
        CStringT_003Cwchar_t_002CStrTraitMFC_DLL_003Cwchar_t_002CATL_003A_003AChTraitsCRT_003Cwchar_t_003E_0020_003E_0020_003E* ptr3 = global::_003CModule_003E.TFlex_002EToCString(&obj, text);
        System.Runtime.CompilerServices.Unsafe.SkipInit(out UniqueIdentifier uniqueIdentifier);
        int num2;
        try
        {
            System.Runtime.CompilerServices.Unsafe.SkipInit(out _GUID gUID);
            IntPtr ptr4 = new IntPtr(&gUID);
            Marshal.StructureToPtr(parameterId, ptr4, fDeleteOld: false);
            _GUID gUID2 = gUID;
            global::_003CModule_003E.UniqueIdentifier_002E_007Bctor_007D(&uniqueIdentifier, &gUID2);
            num2 = global::_003CModule_003E.ps_002ERowElement_002EGetIndex((ps.RowElement*)ptr);
        }
        catch
        {
            //try-fault
            global::_003CModule_003E.___CxxCallUnwindDtor((delegate*<void*, void>)(delegate*<CStringT_003Cwchar_t_002CStrTraitMFC_DLL_003Cwchar_t_002CATL_003A_003AChTraitsCRT_003Cwchar_t_003E_0020_003E_0020_003E*, void>)(&global::_003CModule_003E.ATL_002ECStringT_003Cwchar_t_002CStrTraitMFC_DLL_003Cwchar_t_002CATL_003A_003AChTraitsCRT_003Cwchar_t_003E_0020_003E_0020_003E_002E_007Bdtor_007D), pThis);
            throw;
        }

        global::_003CModule_003E.ps_002EStructureManager_002EChangeRowElemParamDepValue(ptr2, num2, &uniqueIdentifier, ptr3, false, -1, false, -1);
        GC.KeepAlive(this);
    }

    internal unsafe Variable GetCellVariable(Guid parameterId)
    {
        CTFObject* ptr = base.repGet;
        DataTable* dataTable = GetDataTable(bForSet: false);
        System.Runtime.CompilerServices.Unsafe.SkipInit(out _GUID gUID);
        IntPtr ptr2 = new IntPtr(&gUID);
        Marshal.StructureToPtr(parameterId, ptr2, fDeleteOld: false);
        _GUID gUID2 = gUID;
        System.Runtime.CompilerServices.Unsafe.SkipInit(out UniqueIdentifier uniqueIdentifier);
        global::_003CModule_003E.UniqueIdentifier_002E_007Bctor_007D(&uniqueIdentifier, &gUID2);
        UniqueIdentifier* ptr3 = global::_003CModule_003E.ps_002ERowElement_002EGetUID((ps.RowElement*)ptr);
        int num = global::_003CModule_003E.ps_002EDataTable_002EGetVariable(dataTable, ptr3, &uniqueIdentifier);
        if (num >= 0)
        {
            ModelObject result = ModelObject.FromHandle((IntPtr)global::_003CModule_003E.TFDocument_002EGetSafeVar(global::_003CModule_003E.CTfw32Doc_002ECont2D(global::_003CModule_003E.CTFObject_002EGetDocument(ptr)), num));
            GC.KeepAlive(this);
            return (Variable)result;
        }

        GC.KeepAlive(this);
        return null;
    }

    internal unsafe void SetCellVariable(Guid parameterId, Variable variable)
    {
        int num = variable?.VarIndex ?? (-1);
        _ = ProductStructure.repSet;
        CTFObject* ptr = base.repGet;
        System.Runtime.CompilerServices.Unsafe.SkipInit(out StructureManager structureManager);
        StructureManager* intPtr = global::_003CModule_003E.ps_002EStructureManager_002E_007Bctor_007D(&structureManager, global::_003CModule_003E.CTFObject_002EGetDocumentPtr(ptr), true);
        System.Runtime.CompilerServices.Unsafe.SkipInit(out _GUID gUID);
        IntPtr ptr2 = new IntPtr(&gUID);
        Marshal.StructureToPtr(parameterId, ptr2, fDeleteOld: false);
        _GUID gUID2 = gUID;
        System.Runtime.CompilerServices.Unsafe.SkipInit(out UniqueIdentifier uniqueIdentifier);
        global::_003CModule_003E.UniqueIdentifier_002E_007Bctor_007D(&uniqueIdentifier, &gUID2);
        int num2 = global::_003CModule_003E.ps_002ERowElement_002EGetIndex((ps.RowElement*)ptr);
        global::_003CModule_003E.ps_002EStructureManager_002ESetVariable(intPtr, num2, &uniqueIdentifier, num);
        GC.KeepAlive(this);
    }

    internal unsafe Unit GetCellUnit(Guid paramID)
    {
        CTFObject* ptr = base.repGet;
        DataTable* dataTable = GetDataTable(bForSet: false);
        System.Runtime.CompilerServices.Unsafe.SkipInit(out _GUID gUID);
        IntPtr ptr2 = new IntPtr(&gUID);
        Marshal.StructureToPtr(paramID, ptr2, fDeleteOld: false);
        _GUID gUID2 = gUID;
        System.Runtime.CompilerServices.Unsafe.SkipInit(out UniqueIdentifier uniqueIdentifier);
        global::_003CModule_003E.UniqueIdentifier_002E_007Bctor_007D(&uniqueIdentifier, &gUID2);
        UniqueIdentifier* ptr3 = global::_003CModule_003E.ps_002ERowElement_002EGetUID((ps.RowElement*)ptr);
        int num = global::_003CModule_003E.ps_002EDataTable_002EGetUnit(dataTable, ptr3, &uniqueIdentifier);
        if (num >= 0)
        {
            GC.KeepAlive(this);
            return Application.Units.FindUnit(num);
        }

        GC.KeepAlive(this);
        return null;
    }

    internal unsafe void SetCellUnit(Guid paramID, Unit unit)
    {
        int num = ((!(unit != null)) ? (-1) : global::_003CModule_003E.tftypes_002EUnit_002EGetIndex(unit.GetNativeObj()));
        _ = ProductStructure.repSet;
        CTFObject* ptr = base.repGet;
        System.Runtime.CompilerServices.Unsafe.SkipInit(out StructureManager structureManager);
        StructureManager* intPtr = global::_003CModule_003E.ps_002EStructureManager_002E_007Bctor_007D(&structureManager, global::_003CModule_003E.CTFObject_002EGetDocumentPtr(ptr), true);
        System.Runtime.CompilerServices.Unsafe.SkipInit(out _GUID gUID);
        IntPtr ptr2 = new IntPtr(&gUID);
        Marshal.StructureToPtr(paramID, ptr2, fDeleteOld: false);
        _GUID gUID2 = gUID;
        System.Runtime.CompilerServices.Unsafe.SkipInit(out UniqueIdentifier uniqueIdentifier);
        global::_003CModule_003E.UniqueIdentifier_002E_007Bctor_007D(&uniqueIdentifier, &gUID2);
        int num2 = global::_003CModule_003E.ps_002ERowElement_002EGetIndex((ps.RowElement*)ptr);
        global::_003CModule_003E.ps_002EStructureManager_002ESetUnit(intPtr, num2, &uniqueIdentifier, num);
        GC.KeepAlive(this);
    }

    [return: MarshalAs(UnmanagedType.U1)]
    internal unsafe bool IsCellFlagSet(Guid parameterId, CellFlags flag)
    {
        CTFObject* ptr = base.repGet;
        DataTable* dataTable = GetDataTable(bForSet: false);
        System.Runtime.CompilerServices.Unsafe.SkipInit(out _GUID gUID);
        IntPtr ptr2 = new IntPtr(&gUID);
        Marshal.StructureToPtr(parameterId, ptr2, fDeleteOld: false);
        _GUID gUID2 = gUID;
        System.Runtime.CompilerServices.Unsafe.SkipInit(out UniqueIdentifier uniqueIdentifier);
        global::_003CModule_003E.UniqueIdentifier_002E_007Bctor_007D(&uniqueIdentifier, &gUID2);
        UniqueIdentifier* ptr3 = global::_003CModule_003E.ps_002ERowElement_002EGetUID((ps.RowElement*)ptr);
        bool result = global::_003CModule_003E.ps_002EDataTable_002EIsCellFlagSet(dataTable, ptr3, &uniqueIdentifier, flag);
        GC.KeepAlive(this);
        return result;
    }

    internal unsafe void SetCellFlag(Guid parameterId, CellFlags flag, [MarshalAs(UnmanagedType.U1)] bool value)
    {
        CTFObject* ptr = base.repGet;
        _ = ProductStructure.repSet;
        System.Runtime.CompilerServices.Unsafe.SkipInit(out StructureManager structureManager);
        StructureManager* intPtr = global::_003CModule_003E.ps_002EStructureManager_002E_007Bctor_007D(&structureManager, global::_003CModule_003E.CTFObject_002EGetDocumentPtr(ptr), true);
        System.Runtime.CompilerServices.Unsafe.SkipInit(out _GUID gUID);
        IntPtr ptr2 = new IntPtr(&gUID);
        Marshal.StructureToPtr(parameterId, ptr2, fDeleteOld: false);
        _GUID gUID2 = gUID;
        System.Runtime.CompilerServices.Unsafe.SkipInit(out UniqueIdentifier uniqueIdentifier);
        global::_003CModule_003E.UniqueIdentifier_002E_007Bctor_007D(&uniqueIdentifier, &gUID2);
        int num = global::_003CModule_003E.ps_002ERowElement_002EGetIndex((ps.RowElement*)ptr);
        global::_003CModule_003E.ps_002EStructureManager_002ESetFlagForCell(intPtr, num, &uniqueIdentifier, flag, value);
        GC.KeepAlive(this);
    }

    private void throwIfUndoOpen()
    {
    }

    internal unsafe ps.RowElement* GetRowElem([MarshalAs(UnmanagedType.U1)] bool bForSet)
    {
        return (ps.RowElement*)(bForSet ? base.repSet : base.repGet);
    }

    internal unsafe ps.ProductStructure* GetPS([MarshalAs(UnmanagedType.U1)] bool bForSet)
    {
        ProductStructure productStructure = ProductStructure;
        return (ps.ProductStructure*)(bForSet ? productStructure.repSet : productStructure.repGet);
    }

    internal unsafe DataTable* GetDataTable([MarshalAs(UnmanagedType.U1)] bool bForSet)
    {
        ProductStructure productStructure = ProductStructure;
        return global::_003CModule_003E.ps_002EProductStructure_002EGetDataTable((ps.ProductStructure*)(bForSet ? productStructure.repSet : productStructure.repGet));
    }
}
#if false // Журнал декомпиляции
Элементов в кэше: "13"
------------------
Разрешить: "mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"
Найдена одна сборка: "mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"
Загрузить из: "C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.7.2\mscorlib.dll"
------------------
Разрешить: "PresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35"
Не удалось найти по имени: "PresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35"
------------------
Разрешить: "PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35"
Не удалось найти по имени: "PresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35"
------------------
Разрешить: "System.ComponentModel.DataAnnotations, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35"
Не удалось найти по имени: "System.ComponentModel.DataAnnotations, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35"
------------------
Разрешить: "System.Core, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"
Найдена одна сборка: "System.Core, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"
Загрузить из: "C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.7.2\System.Core.dll"
------------------
Разрешить: "System, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"
Найдена одна сборка: "System, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"
Загрузить из: "C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.7.2\System.dll"
------------------
Разрешить: "System.Drawing, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"
Найдена одна сборка: "System.Drawing, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"
Загрузить из: "C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.7.2\System.Drawing.dll"
------------------
Разрешить: "System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"
Найдена одна сборка: "System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"
Загрузить из: "C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.7.2\System.Windows.Forms.dll"
------------------
Разрешить: "TFlex.CAD.Strings, Version=17.1.30.0, Culture=neutral, PublicKeyToken=c2a83ea67ea56687"
Не удалось найти по имени: "TFlex.CAD.Strings, Version=17.1.30.0, Culture=neutral, PublicKeyToken=c2a83ea67ea56687"
------------------
Разрешить: "TFlexAPIData, Version=17.1.30.0, Culture=neutral, PublicKeyToken=effde555051a6517"
Найдена одна сборка: "TFlexAPIData, Version=17.1.30.0, Culture=neutral, PublicKeyToken=effde555051a6517"
Загрузить из: "C:\Users\Lerik\YandexDisk\templates\!Programming\Git\Собиратель\SpecCollector\bin\Debug\TFlexAPIData.dll"
------------------
Разрешить: "WindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35"
Не удалось найти по имени: "WindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35"
#endif

