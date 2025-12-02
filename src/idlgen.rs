use super::error::Error;
use windows::{
    Win32::System::{
        Com::{
            CoInitialize, FUNCDESC, IMPLTYPEFLAGS, INVOKE_PROPERTYGET, INVOKE_PROPERTYPUT,
            INVOKE_PROPERTYPUTREF, INVOKEKIND, ITypeInfo, ITypeInfo2, ITypeLib, ITypeLib2,
            TKIND_ALIAS, TKIND_COCLASS, TKIND_DISPATCH, TKIND_ENUM, TKIND_INTERFACE, TKIND_MODULE,
            TKIND_RECORD, TYPEATTR, TYPEDESC, VARDESC,
        },
        Ole::LoadTypeLib,
        Variant::{
            VARIANT, VT_BOOL, VT_BSTR, VT_CY, VT_DATE, VT_DECIMAL, VT_DISPATCH, VT_EMPTY, VT_ERROR,
            VT_HRESULT, VT_I1, VT_I2, VT_I4, VT_I8, VT_INT, VT_LPSTR, VT_LPWSTR, VT_NULL, VT_PTR,
            VT_R4, VT_R8, VT_SAFEARRAY, VT_UI1, VT_UI2, VT_UI4, VT_UI8, VT_UINT, VT_UNKNOWN,
            VT_USERDEFINED, VT_VARIANT, VT_VOID,
        },
    },
    core::{Interface, PCWSTR},
};
use windows_core::{BSTR, HSTRING};

#[derive(Debug)]
pub struct BuildResult {
    pub num_missing_types: usize,
    pub num_types_not_found: usize,
}

struct TypeLibInfo {
    tlib: Option<ITypeLib>,
}

impl TypeLibInfo {
    fn new() -> Self {
        TypeLibInfo { tlib: None }
    }

    fn load_type_lib(&mut self, path: &std::path::Path) -> Result<(), Error> {
        unsafe {
            let _ = CoInitialize(None);
        }
        let path_str = path.to_str().ok_or(Error::TypeLibNotLoaded)?;
        let path_hstring = HSTRING::from(path_str);
        let path_pcwstr = PCWSTR::from_raw(path_hstring.as_ptr());

        let type_lib = unsafe { LoadTypeLib(path_pcwstr) }?;
        self.tlib = Some(type_lib);
        Ok(())
    }

    fn get_lib_attr(&self) -> Result<*mut windows::Win32::System::Com::TLIBATTR, Error> {
        if let Some(tlib) = &self.tlib {
            unsafe { Ok(tlib.GetLibAttr()?) }
        } else {
            Err(Error::TypeLibNotLoaded)
        }
    }

    fn get_documentation(&self, index: i32) -> Result<(BSTR, BSTR), Error> {
        if let Some(tlib) = &self.tlib {
            let mut name = BSTR::new();
            let mut doc_string = BSTR::new();
            let mut help_context = 0;
            unsafe {
                tlib.GetDocumentation(
                    index,
                    Some(&mut name),
                    Some(&mut doc_string),
                    &mut help_context,
                    None,
                )?
            };
            Ok((name, doc_string))
        } else {
            Err(Error::TypeLibNotLoaded)
        }
    }

    fn get_type_info_count(&self) -> u32 {
        if let Some(tlib) = &self.tlib {
            unsafe { tlib.GetTypeInfoCount() }
        } else {
            0
        }
    }

    fn get_type_info(&self, index: u32) -> Result<ITypeInfo, Error> {
        if let Some(tlib) = &self.tlib {
            unsafe { Ok(tlib.GetTypeInfo(index)?) }
        } else {
            Err(Error::TypeLibNotLoaded)
        }
    }
}

pub fn get_library_name(tlb_path: &std::path::Path) -> Result<String, Error> {
    let mut type_lib_info = TypeLibInfo::new();
    type_lib_info.load_type_lib(tlb_path)?;
    let (name, _) = type_lib_info.get_documentation(-1)?;
    Ok(name.to_string())
}

pub fn build_tlb<W>(tlb_path: &std::path::Path, mut out: W) -> Result<BuildResult, Error>
where
    W: std::io::Write,
{
    let mut type_lib_info = TypeLibInfo::new();
    type_lib_info.load_type_lib(tlb_path)?;

    let lib_attr = unsafe { &*type_lib_info.get_lib_attr()? };
    let (name, doc_string) = type_lib_info.get_documentation(-1)?;

    writeln!(out, "// Decompilado desde {}", tlb_path.display())?;
    writeln!(out, "[")?;
    writeln!(out, "  uuid({:?}),", lib_attr.guid)?;
    writeln!(
        out,
        "  version({}.{}),",
        lib_attr.wMajorVerNum, lib_attr.wMinorVerNum
    )?;
    writeln!(out, "  helpstring(\"{}\"),", doc_string)?;

    if let Some(tlib) = &type_lib_info.tlib {
        if let Ok(tlib2) = tlib.cast::<ITypeLib2>() {
            unsafe {
                print_lib_custom_data(&tlib2, &mut out)?;
            }
        }
    }

    writeln!(out, "]")?;
    writeln!(out, "library {}", name)?;
    writeln!(out, "{{")?;

    // Standard imports often found in IDLs
    writeln!(out, "    importlib(\"stdole2.tlb\");")?;
    writeln!(out, "")?;

    let count = type_lib_info.get_type_info_count();

    // Forward declarations
    for i in 0..count {
        if let Ok(type_info) = type_lib_info.get_type_info(i) {
            unsafe {
                let type_attr: *mut TYPEATTR = type_info.GetTypeAttr()?;
                let type_kind = (*type_attr).typekind;
                let type_flags = (*type_attr).wTypeFlags;
                let (name, _) = get_type_documentation(&type_info, -1);

                match type_kind {
                    TKIND_INTERFACE => {
                        writeln!(out, "    interface {};", name)?;
                    }
                    TKIND_DISPATCH => {
                        let is_dual = (type_flags & 0x40) != 0; // TYPEFLAG_FDUAL
                        if is_dual {
                            writeln!(out, "    interface {};", name)?;
                        } else {
                            writeln!(out, "    dispinterface {};", name)?;
                        }
                    }
                    TKIND_COCLASS => {
                        writeln!(out, "    coclass {};", name)?;
                    }
                    _ => {}
                }
                type_info.ReleaseTypeAttr(type_attr);
            }
        }
    }
    writeln!(out, "")?;

    // Print enums first
    let count = type_lib_info.get_type_info_count();
    for i in 0..count {
        if let Ok(type_info) = type_lib_info.get_type_info(i) {
            unsafe {
                let type_attr = type_info.GetTypeAttr()?;
                if (*type_attr).typekind == TKIND_ENUM {
                    print_type_info(&type_info, &mut out)?;
                }
                type_info.ReleaseTypeAttr(type_attr);
            }
        }
    }
    writeln!(out, "")?;

    let count = type_lib_info.get_type_info_count();
    for i in 0..count {
        if let Ok(type_info) = type_lib_info.get_type_info(i) {
            unsafe {
                let type_attr = type_info.GetTypeAttr()?;
                if (*type_attr).typekind != TKIND_ENUM {
                    print_type_info(&type_info, &mut out)?;
                }
                type_info.ReleaseTypeAttr(type_attr);
            }
        }
    }

    writeln!(out, "}};")?;

    Ok(BuildResult {
        num_missing_types: 0,
        num_types_not_found: 0,
    })
}

fn print_type_info<W>(type_info: &ITypeInfo, out: &mut W) -> Result<(), Error>
where
    W: std::io::Write,
{
    unsafe {
        let type_attr: *mut TYPEATTR = type_info.GetTypeAttr()?;
        let type_kind = (*type_attr).typekind;
        let guid = (*type_attr).guid;
        let (name, doc_string) = get_type_documentation(type_info, -1);
        let type_flags = (*type_attr).wTypeFlags;

        writeln!(out, "    [")?;
        writeln!(out, "      odl,")?;
        writeln!(out, "      uuid({:?}),", guid)?;
        if type_kind == TKIND_INTERFACE || type_kind == TKIND_DISPATCH || type_kind == TKIND_COCLASS
        {
            writeln!(
                out,
                "      version({}.{}),",
                (*type_attr).wMajorVerNum,
                (*type_attr).wMinorVerNum
            )?;
        }
        if !doc_string.is_empty() {
            writeln!(out, "      helpstring(\"{}\"),", doc_string)?;
        }

        // Check attributes
        if (type_flags & 0x40) != 0 {
            writeln!(out, "      dual,")?;
        } // TYPEFLAG_FDUAL
        if (type_flags & 0x20) != 0 {
            writeln!(out, "      oleautomation,")?;
        } // TYPEFLAG_FOLEAUTOMATION
        if (type_flags & 0x80) != 0 {
            writeln!(out, "      nonextensible,")?;
        } // TYPEFLAG_FNONEXTENSIBLE

        // Custom attributes
        if let Ok(type_info2) = type_info.cast::<ITypeInfo2>() {
            print_custom_data(&type_info2, out)?;
        }

        writeln!(out, "    ]")?;

        match type_kind {
            TKIND_INTERFACE => {
                // Find base interface
                let mut base_name = String::new();
                if (*type_attr).cImplTypes > 0 {
                    if let Ok(href) = type_info.GetRefTypeOfImplType(0) {
                        if let Ok(base_info) = type_info.GetRefTypeInfo(href) {
                            base_name = get_name(&base_info);
                        }
                    }
                }

                if !base_name.is_empty() {
                    writeln!(out, "    interface {} : {} {{", name, base_name)?;
                } else {
                    writeln!(out, "    interface {} {{", name)?;
                }

                // Print properties and methods
                for i in 0..(*type_attr).cFuncs {
                    if let Ok(func_desc) = type_info.GetFuncDesc(i as u32) {
                        print_function(type_info, &*func_desc, out)?;
                        type_info.ReleaseFuncDesc(func_desc);
                    }
                }

                writeln!(out, "    }};")?;
            }
            TKIND_DISPATCH => {
                let is_dual = (type_flags & 0x40) != 0;
                if is_dual {
                    // Dual interface, treat as standard interface
                    // Find base interface
                    let mut base_name = String::new();
                    if (*type_attr).cImplTypes > 0 {
                        if let Ok(href) = type_info.GetRefTypeOfImplType(0) {
                            if let Ok(base_info) = type_info.GetRefTypeInfo(href) {
                                base_name = get_name(&base_info);
                            }
                        }
                    }

                    if !base_name.is_empty() {
                        writeln!(out, "    interface {} : {} {{", name, base_name)?;
                    } else {
                        writeln!(out, "    interface {} {{", name)?;
                    }

                    // Print properties and methods
                    for i in 0..(*type_attr).cFuncs {
                        if let Ok(func_desc) = type_info.GetFuncDesc(i as u32) {
                            print_function(type_info, &*func_desc, out)?;
                            type_info.ReleaseFuncDesc(func_desc);
                        }
                    }
                    writeln!(out, "    }};")?;
                } else {
                    // Pure dispinterface
                    writeln!(out, "    dispinterface {} {{", name)?;

                    if (*type_attr).cVars > 0 {
                        writeln!(out, "    properties:")?;
                        for i in 0..(*type_attr).cVars {
                            if let Ok(var_desc) = type_info.GetVarDesc(i as u32) {
                                print_disp_property(type_info, &*var_desc, out)?;
                                type_info.ReleaseVarDesc(var_desc);
                            }
                        }
                    }

                    if (*type_attr).cFuncs > 0 {
                        writeln!(out, "    methods:")?;
                        for i in 0..(*type_attr).cFuncs {
                            if let Ok(func_desc) = type_info.GetFuncDesc(i as u32) {
                                print_function(type_info, &*func_desc, out)?;
                                type_info.ReleaseFuncDesc(func_desc);
                            }
                        }
                    }
                    writeln!(out, "    }};")?;
                }
            }
            TKIND_ENUM => {
                writeln!(out, "    enum {} {{", name)?;
                for i in 0..(*type_attr).cVars {
                    if let Ok(var_desc) = type_info.GetVarDesc(i as u32) {
                        print_var(type_info, &*var_desc, out)?;
                        type_info.ReleaseVarDesc(var_desc);
                    }
                }
                writeln!(out, "    }};")?;
            }
            TKIND_COCLASS => {
                writeln!(out, "    coclass {} {{", name)?;
                for i in 0..(*type_attr).cImplTypes {
                    if let Ok(href) = type_info.GetRefTypeOfImplType(i as u32) {
                        if let Ok(ref_type_info) = type_info.GetRefTypeInfo(href) {
                            let ref_name = get_name(&ref_type_info);
                            let impl_flags = type_info
                                .GetImplTypeFlags(i as u32)
                                .unwrap_or(IMPLTYPEFLAGS::default());
                            // Check for [default]
                            let default_str = if (impl_flags.0 & 1) != 0 {
                                "[default] "
                            } else {
                                ""
                            };
                            // Check for [source] (2)
                            let source_str = if (impl_flags.0 & 2) != 0 {
                                "[source] "
                            } else {
                                ""
                            };

                            writeln!(
                                out,
                                "        {}{}{} {};",
                                default_str, source_str, "interface", ref_name
                            )?;
                        }
                    }
                }
                writeln!(out, "    }};")?;
            }
            TKIND_ALIAS => {
                // typedef
                if let Ok(_) = type_info
                    .GetRefTypeOfImplType(0)
                    .and_then(|href| type_info.GetRefTypeInfo(href))
                {
                    // Simplified alias handling
                    writeln!(out, "    typedef {};", name)?;
                }
            }
            TKIND_RECORD => {
                writeln!(out, "    typedef struct tag{} {{", name)?;
                for i in 0..(*type_attr).cVars {
                    if let Ok(var_desc) = type_info.GetVarDesc(i as u32) {
                        print_record_member(type_info, &*var_desc, out)?;
                        type_info.ReleaseVarDesc(var_desc);
                    }
                }
                writeln!(out, "    }} {};", name)?;
            }
            TKIND_MODULE => {
                // Try to get dllname
                let mut dll_name = String::new();
                if (*type_attr).cFuncs > 0 {
                    if let Ok(func_desc) = type_info.GetFuncDesc(0) {
                        if let Ok(dll) =
                            get_dll_entry(type_info, (*func_desc).memid, (*func_desc).invkind)
                        {
                            dll_name = dll;
                        }
                        type_info.ReleaseFuncDesc(func_desc);
                    }
                }

                writeln!(out, "    [")?;
                if !dll_name.is_empty() {
                    writeln!(out, "      dllname(\"{}\"),", dll_name)?;
                }
                writeln!(out, "      uuid({:?}),", guid)?;
                if !doc_string.is_empty() {
                    writeln!(out, "      helpstring(\"{}\"),", doc_string)?;
                }
                writeln!(out, "    ]")?;
                writeln!(out, "    module {} {{", name)?;

                // Print vars (constants)
                for i in 0..(*type_attr).cVars {
                    if let Ok(var_desc) = type_info.GetVarDesc(i as u32) {
                        print_module_const(type_info, &*var_desc, out)?;
                        type_info.ReleaseVarDesc(var_desc);
                    }
                }

                // Print funcs
                for i in 0..(*type_attr).cFuncs {
                    if let Ok(func_desc) = type_info.GetFuncDesc(i as u32) {
                        type_info.ReleaseFuncDesc(func_desc);
                    }
                }
                writeln!(out, "    }};")?;
            }
            _ => {
                writeln!(out, "    // Unsupported type kind: {:?}", type_kind)?;
            }
        }
        writeln!(out, "")?;

        type_info.ReleaseTypeAttr(type_attr);
    }
    Ok(())
}

unsafe fn print_module_const<W>(
    type_info: &ITypeInfo,
    var_desc: &VARDESC,
    out: &mut W,
) -> Result<(), Error>
where
    W: std::io::Write,
{
    let memid = var_desc.memid;
    let (name, _) = unsafe { get_type_documentation(type_info, memid) };

    if let Some(val) = unsafe { var_desc.Anonymous.lpvarValue.as_ref() } {
        // Assuming int/long for now as per Olewoo example
        let val_int = unsafe { val.Anonymous.Anonymous.Anonymous.lVal };
        // Handle negative hex printing if needed, but simple print for now
        writeln!(out, "        const int {} = {};", name, val_int)?;
    }
    Ok(())
}

unsafe fn get_dll_entry(
    type_info: &ITypeInfo,
    memid: i32,
    invkind: INVOKEKIND,
) -> Result<String, Error> {
    let mut dll_name = BSTR::new();
    let mut name = BSTR::new();
    let mut ordinal = 0u16;
    unsafe {
        type_info.GetDllEntry(
            memid,
            invkind,
            Some(&mut dll_name as *mut BSTR),
            Some(&mut name as *mut BSTR),
            &mut ordinal,
        )?;
    }
    Ok(dll_name.to_string())
}

unsafe fn print_lib_custom_data<W>(type_lib2: &ITypeLib2, out: &mut W) -> Result<(), Error>
where
    W: std::io::Write,
{
    let cust_data = unsafe { type_lib2.GetAllCustData()? };

    for i in 0..cust_data.cCustData {
        let item = unsafe { &*cust_data.prgCustData.offset(i as isize) };
        let guid = item.guid;
        let val = &item.varValue;

        let vt = unsafe { val.Anonymous.Anonymous.vt };
        if vt == VT_BSTR {
            let bstr_val = unsafe { &val.Anonymous.Anonymous.Anonymous.bstrVal };
            let s = bstr_val.to_string();
            writeln!(out, "  custom({:?}, \"{}\"),", guid, s)?;
        }
    }
    Ok(())
}

unsafe fn print_custom_data<W>(type_info2: &ITypeInfo2, out: &mut W) -> Result<(), Error>
where
    W: std::io::Write,
{
    let cust_data = unsafe { type_info2.GetAllCustData()? };

    for i in 0..cust_data.cCustData {
        let item = unsafe { &*cust_data.prgCustData.offset(i as isize) };
        let guid = item.guid;
        let val = &item.varValue;

        let vt = unsafe { val.Anonymous.Anonymous.vt };
        if vt == VT_BSTR {
            let bstr_val = unsafe { &val.Anonymous.Anonymous.Anonymous.bstrVal };
            // bstr_val is ManuallyDrop<BSTR>
            let s = bstr_val.to_string();
            writeln!(out, "      custom({:?}, \"{}\"),", guid, s)?;
        }
    }
    Ok(())
}

unsafe fn print_function<W>(
    type_info: &ITypeInfo,
    func_desc: &FUNCDESC,
    out: &mut W,
) -> Result<(), Error>
where
    W: std::io::Write,
{
    let memid = func_desc.memid;

    // Filter IUnknown and IDispatch methods
    // Olewoo filters 0x60000000 to 0x60020000
    if memid >= 0x60000000 && memid < 0x60020000 {
        return Ok(());
    }

    let (name, _) = unsafe { get_type_documentation(type_info, memid) };

    let invoke_kind = func_desc.invkind;
    let prop_prefix = match invoke_kind {
        INVOKE_PROPERTYGET => "[propget] ",
        INVOKE_PROPERTYPUT => "[propput] ",
        INVOKE_PROPERTYPUTREF => "[propputref] ",
        _ => "",
    };

    let ret_type = unsafe { type_desc_to_string(type_info, &func_desc.elemdescFunc.tdesc) };

    writeln!(out, "        [id(0x{:08x})]", memid)?;
    write!(out, "        {}HRESULT {} (", prop_prefix, name)?;

    // Get parameter names
    let mut names: Vec<BSTR> = vec![BSTR::new(); (func_desc.cParams + 1) as usize];
    let mut c_names = 0;
    unsafe {
        type_info
            .GetNames(memid, names.as_mut_slice(), &mut c_names)
            .ok();
    }
    // names[0] is the function name, names[1..] are params

    for i in 0..func_desc.cParams {
        let elem_desc = unsafe { *func_desc.lprgelemdescParam.offset(i as isize) };
        let param_type = unsafe { type_desc_to_string(type_info, &elem_desc.tdesc) };

        // Get param name
        let param_name = if (i + 1) < c_names as i16 {
            names[(i + 1) as usize].to_string()
        } else {
            format!("arg{}", i)
        };

        // Param attributes
        let param_flags = unsafe { elem_desc.Anonymous.paramdesc.wParamFlags };
        let mut attrs: Vec<String> = Vec::new();
        if (param_flags.0 & 1) != 0 {
            attrs.push("in".to_string());
        } // PARAMFLAG_FIN
        if (param_flags.0 & 2) != 0 {
            attrs.push("out".to_string());
        } // PARAMFLAG_FOUT
        if (param_flags.0 & 4) != 0 {
            attrs.push("lcid".to_string());
        } // PARAMFLAG_FLCID
        if (param_flags.0 & 8) != 0 {
            attrs.push("retval".to_string());
        } // PARAMFLAG_FRETVAL
        if (param_flags.0 & 16) != 0 {
            attrs.push("optional".to_string());
        } // PARAMFLAG_FOPT
        if (param_flags.0 & 32) != 0 {
            let default_val = unsafe {
                let param_desc_ex = elem_desc.Anonymous.paramdesc.pparamdescex;
                if !param_desc_ex.is_null() {
                    let variant = &(*param_desc_ex).varDefaultValue;
                    variant_to_string(variant)
                } else {
                    String::new()
                }
            };
            if !default_val.is_empty() {
                attrs.push(format!("defaultvalue({})", default_val));
            } else {
                attrs.push("defaultvalue".to_string());
            }
        } // PARAMFLAG_FHASDEFAULT

        let attr_str = if !attrs.is_empty() {
            format!("[{}] ", attrs.join(", "))
        } else {
            String::new()
        };

        if i > 0 {
            write!(out, ", ")?;
        }
        write!(out, "{}{} {}", attr_str, param_type, param_name)?;
    }

    // Handle return value for property get or functions returning values
    if invoke_kind == INVOKE_PROPERTYGET || ret_type != "void" {
        if func_desc.cParams > 0 {
            write!(out, ", ")?;
        }

        let mut ret_name = "val";
        for i in 1..c_names {
            if names[i as usize].to_string().eq_ignore_ascii_case("val") {
                ret_name = "retVal";
                break;
            }
        }

        write!(out, "[out, retval] {}* {}", ret_type, ret_name)?;
    }

    writeln!(out, ");")?;
    Ok(())
}

unsafe fn print_disp_property<W>(
    type_info: &ITypeInfo,
    var_desc: &VARDESC,
    out: &mut W,
) -> Result<(), Error>
where
    W: std::io::Write,
{
    let memid = var_desc.memid;
    let (name, _) = unsafe { get_type_documentation(type_info, memid) };
    let type_name = unsafe { type_desc_to_string(type_info, &var_desc.elemdescVar.tdesc) };

    writeln!(out, "        [id(0x{:08x})] {} {};", memid, type_name, name)?;
    Ok(())
}

unsafe fn print_var<W>(type_info: &ITypeInfo, var_desc: &VARDESC, out: &mut W) -> Result<(), Error>
where
    W: std::io::Write,
{
    // For enums
    let memid = var_desc.memid;
    let (name, _) = unsafe { get_type_documentation(type_info, memid) };

    if let Some(val) = unsafe { var_desc.Anonymous.lpvarValue.as_ref() } {
        // This is tricky without full Variant support, assuming I4 for enums usually
        let val_int = unsafe { val.Anonymous.Anonymous.Anonymous.lVal };
        writeln!(out, "        {} = {},", name, val_int)?;
    } else {
        writeln!(out, "        {},", name)?;
    }
    Ok(())
}

unsafe fn print_record_member<W>(
    type_info: &ITypeInfo,
    var_desc: &VARDESC,
    out: &mut W,
) -> Result<(), Error>
where
    W: std::io::Write,
{
    let memid = var_desc.memid;
    let (name, _) = unsafe { get_type_documentation(type_info, memid) };
    let type_name = unsafe { type_desc_to_string(type_info, &var_desc.elemdescVar.tdesc) };
    writeln!(out, "        {} {};", type_name, name)?;
    Ok(())
}

unsafe fn get_type_documentation(type_info: &ITypeInfo, index: i32) -> (String, String) {
    let mut name = BSTR::new();
    let mut doc_string = BSTR::new();
    unsafe {
        let _ =
            type_info.GetDocumentation(index, Some(&mut name), Some(&mut doc_string), &mut 0, None);
    }
    (name.to_string(), doc_string.to_string())
}

unsafe fn get_name(type_info: &ITypeInfo) -> String {
    let (name, _) = unsafe { get_type_documentation(type_info, -1) };
    name
}

unsafe fn type_desc_to_string(type_info: &ITypeInfo, tdesc: &TYPEDESC) -> String {
    match tdesc.vt {
        VT_I2 => "short".to_string(),
        VT_I4 => "long".to_string(),
        VT_R4 => "float".to_string(),
        VT_R8 => "double".to_string(),
        VT_CY => "CURRENCY".to_string(),
        VT_DATE => "DATE".to_string(),
        VT_BSTR => "BSTR".to_string(),
        VT_DISPATCH => "IDispatch*".to_string(),
        VT_ERROR => "SCODE".to_string(),
        VT_BOOL => "VARIANT_BOOL".to_string(),
        VT_VARIANT => "VARIANT".to_string(),
        VT_UNKNOWN => "IUnknown*".to_string(),
        VT_DECIMAL => "DECIMAL".to_string(),
        VT_I1 => "char".to_string(),
        VT_UI1 => "unsigned char".to_string(),
        VT_UI2 => "unsigned short".to_string(),
        VT_UI4 => "unsigned long".to_string(),
        VT_I8 => "int64".to_string(),
        VT_UI8 => "uint64".to_string(),
        VT_INT => "int".to_string(),
        VT_UINT => "unsigned int".to_string(),
        VT_VOID => "void".to_string(),
        VT_HRESULT => "HRESULT".to_string(),
        VT_PTR => {
            let pointed_type = unsafe { type_desc_to_string(type_info, &*tdesc.Anonymous.lptdesc) };
            format!("{}*", pointed_type)
        }
        VT_SAFEARRAY => "SAFEARRAY".to_string(),
        VT_USERDEFINED => {
            if let Ok(ref_type_info) = unsafe { type_info.GetRefTypeInfo(tdesc.Anonymous.hreftype) }
            {
                let name = unsafe { get_name(&ref_type_info) };
                unsafe {
                    if let Ok(type_attr) = ref_type_info.GetTypeAttr() {
                        let kind = (*type_attr).typekind;
                        ref_type_info.ReleaseTypeAttr(type_attr);
                        if kind == TKIND_ENUM {
                            format!("enum {}", name)
                        } else {
                            name
                        }
                    } else {
                        name
                    }
                }
            } else {
                "UnknownUserDefined".to_string()
            }
        }
        VT_LPSTR => "LPSTR".to_string(),
        VT_LPWSTR => "LPWSTR".to_string(),
        _ => format!("TYPE_{}", tdesc.vt.0),
    }
}

unsafe fn variant_to_string(variant: &VARIANT) -> String {
    let variant = &*variant;
    unsafe {
        match variant.Anonymous.Anonymous.vt {
            VT_I2 => variant.Anonymous.Anonymous.Anonymous.iVal.to_string(),
            VT_I4 => variant.Anonymous.Anonymous.Anonymous.lVal.to_string(),
            VT_R4 => variant.Anonymous.Anonymous.Anonymous.fltVal.to_string(),
            VT_R8 => variant.Anonymous.Anonymous.Anonymous.dblVal.to_string(),
            VT_BOOL => {
                if variant.Anonymous.Anonymous.Anonymous.boolVal.as_bool() {
                    "TRUE".to_string()
                } else {
                    "FALSE".to_string()
                }
            }
            VT_BSTR => {
                let bstr = &variant.Anonymous.Anonymous.Anonymous.bstrVal;
                if !bstr.is_empty() {
                    format!("\"{}\"", bstr.to_string())
                } else {
                    "\"\"".to_string()
                }
            }
            VT_EMPTY => "".to_string(),
            VT_NULL => "null".to_string(),
            _ => format!("/* vt: {} */", variant.Anonymous.Anonymous.vt.0),
        }
    }
}
