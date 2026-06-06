use super::error::Error;
use crate::flags::{ImplTypeFlags, ParamFlags, TypeFlags};
use windows::{
    Win32::System::{
        Com::{
            CoInitialize, FUNCDESC, IDispatch, IMPLTYPEFLAGS, INVOKE_PROPERTYGET,
            INVOKE_PROPERTYPUT, INVOKE_PROPERTYPUTREF, INVOKEKIND, ITypeInfo, ITypeInfo2, ITypeLib,
            ITypeLib2, TKIND_ALIAS, TKIND_COCLASS, TKIND_DISPATCH, TKIND_ENUM, TKIND_INTERFACE,
            TKIND_MODULE, TKIND_RECORD, TKIND_UNION, TLIBATTR, TYPEATTR, TYPEDESC, VARDESC,
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
use windows_core::{BSTR, HSTRING, IUnknown};

// RAII Wrappers for COM structures

struct SafeTypeAttr<'a> {
    type_info: &'a ITypeInfo,
    attr: *mut TYPEATTR,
}

impl<'a> SafeTypeAttr<'a> {
    fn new(type_info: &'a ITypeInfo) -> Result<Self, Error> {
        unsafe {
            let attr = type_info.GetTypeAttr()?;
            Ok(Self { type_info, attr })
        }
    }
}

impl<'a> std::ops::Deref for SafeTypeAttr<'a> {
    type Target = TYPEATTR;

    fn deref(&self) -> &Self::Target {
        unsafe { &*self.attr }
    }
}

impl<'a> Drop for SafeTypeAttr<'a> {
    fn drop(&mut self) {
        unsafe {
            self.type_info.ReleaseTypeAttr(self.attr);
        }
    }
}

struct SafeFuncDesc<'a> {
    type_info: &'a ITypeInfo,
    desc: *mut FUNCDESC,
}

impl<'a> SafeFuncDesc<'a> {
    fn new(type_info: &'a ITypeInfo, index: u32) -> Result<Self, Error> {
        unsafe {
            let desc = type_info.GetFuncDesc(index)?;
            Ok(Self { type_info, desc })
        }
    }
}

impl<'a> std::ops::Deref for SafeFuncDesc<'a> {
    type Target = FUNCDESC;

    fn deref(&self) -> &Self::Target {
        unsafe { &*self.desc }
    }
}

impl<'a> Drop for SafeFuncDesc<'a> {
    fn drop(&mut self) {
        unsafe {
            self.type_info.ReleaseFuncDesc(self.desc);
        }
    }
}

struct SafeVarDesc<'a> {
    type_info: &'a ITypeInfo,
    desc: *mut VARDESC,
}

impl<'a> SafeVarDesc<'a> {
    fn new(type_info: &'a ITypeInfo, index: u32) -> Result<Self, Error> {
        unsafe {
            let desc = type_info.GetVarDesc(index)?;
            Ok(Self { type_info, desc })
        }
    }
}

impl<'a> std::ops::Deref for SafeVarDesc<'a> {
    type Target = VARDESC;

    fn deref(&self) -> &Self::Target {
        unsafe { &*self.desc }
    }
}

impl<'a> Drop for SafeVarDesc<'a> {
    fn drop(&mut self) {
        unsafe {
            self.type_info.ReleaseVarDesc(self.desc);
        }
    }
}

struct SafeLibAttr<'a> {
    tlib: &'a ITypeLib,
    attr: *mut TLIBATTR,
}

impl<'a> SafeLibAttr<'a> {
    fn new(tlib: &'a ITypeLib) -> Result<Self, Error> {
        unsafe {
            let attr = tlib.GetLibAttr()?;
            Ok(Self { tlib, attr })
        }
    }
}

impl<'a> std::ops::Deref for SafeLibAttr<'a> {
    type Target = TLIBATTR;

    fn deref(&self) -> &Self::Target {
        unsafe { &*self.attr }
    }
}

impl<'a> Drop for SafeLibAttr<'a> {
    fn drop(&mut self) {
        unsafe {
            self.tlib.ReleaseTLibAttr(self.attr);
        }
    }
}

pub struct TypeLibInfo {
    tlib: Option<ITypeLib>,
}

impl TypeLibInfo {
    pub fn new() -> Self {
        TypeLibInfo { tlib: None }
    }

    pub fn load_type_lib(&mut self, path: &std::path::Path) -> Result<(), Error> {
        unsafe {
            let _ = CoInitialize(None);
        }

        let abs_path = path.canonicalize().map_err(|e| Error::IoError(e))?;
        let path_str = abs_path.to_str().ok_or(Error::TypeLibNotLoaded)?;
        let path_hstring = HSTRING::from(path_str);
        let path_pcwstr = PCWSTR::from_raw(path_hstring.as_ptr());

        let type_lib = unsafe { LoadTypeLib(path_pcwstr) }?;
        self.tlib = Some(type_lib);
        Ok(())
    }

    // Helper using RAII wrapper
    #[allow(dead_code)]
    fn get_lib_attr_safe(&self) -> Result<SafeLibAttr<'_>, Error> {
        if let Some(tlib) = &self.tlib {
            SafeLibAttr::new(tlib)
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

    pub fn get_type_info_count(&self) -> u32 {
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

    pub fn get_type_name_and_kind(&self, index: u32) -> Result<(String, String), Error> {
        let type_info = self.get_type_info(index)?;
        let type_attr = SafeTypeAttr::new(&type_info)?;

        let kind = match type_attr.typekind {
            TKIND_ENUM => "Enum",
            TKIND_RECORD => "Record",
            TKIND_MODULE => "Module",
            TKIND_INTERFACE => "Interface",
            TKIND_DISPATCH => "Dispatch",
            TKIND_COCLASS => "CoClass",
            TKIND_ALIAS => "Alias",
            TKIND_UNION => "Union",
            _ => "Unknown",
        }
        .to_string();
        let (name, _) = unsafe { get_type_documentation(&type_info, -1) };
        Ok((name, kind))
    }

    pub fn get_type_idl(&self, index: u32) -> Result<String, Error> {
        let type_info = self.get_type_info(index)?;
        let mut out = Vec::new();
        print_type_info(&type_info, &mut out)?;
        Ok(String::from_utf8_lossy(&out).to_string())
    }

    pub fn get_type_methods(&self, index: u32) -> Result<Vec<MethodInfo>, Error> {
        let type_info = self.get_type_info(index)?;
        let mut methods = Vec::new();
        let type_attr = SafeTypeAttr::new(&type_info)?;

        for i in 0..type_attr.cFuncs {
            if let Ok(func_desc) = SafeFuncDesc::new(&type_info, i as u32) {
                if let Ok(info) = get_function_info(&type_info, &func_desc) {
                    methods.push(info);
                }
            }
        }
        Ok(methods)
    }

    pub fn get_type_enums(&self, index: u32) -> Result<Vec<EnumItemInfo>, Error> {
        let type_info = self.get_type_info(index)?;
        let mut enums = Vec::new();
        let type_attr = SafeTypeAttr::new(&type_info)?;

        if type_attr.typekind == TKIND_ENUM {
            for i in 0..type_attr.cVars {
                if let Ok(var_desc) = SafeVarDesc::new(&type_info, i as u32) {
                    if let Ok(info) = unsafe { get_enum_info(&type_info, &var_desc) } {
                        enums.push(info);
                    }
                }
            }
        }
        Ok(enums)
    }
}

#[derive(Debug, Clone)]
pub struct EnumItemInfo {
    pub name: String,
    pub value: String,
}

#[derive(Debug, Clone)]
pub struct ParamInfo {
    pub name: String,
    pub type_name: String,
    pub flags: Vec<String>,
    pub default_value: Option<String>,
}

#[derive(Debug, Clone)]
pub struct MethodInfo {
    pub name: String,
    pub ret_type: String,
    pub params: Vec<ParamInfo>,
    pub _invoke_kind: String,
}

unsafe fn get_enum_info(type_info: &ITypeInfo, var_desc: &VARDESC) -> Result<EnumItemInfo, Error> {
    let memid = var_desc.memid;
    let (name, _) = unsafe { get_type_documentation(type_info, memid) };

    let value = if let Some(val) = unsafe { var_desc.Anonymous.lpvarValue.as_ref() } {
        unsafe { val.Anonymous.Anonymous.Anonymous.lVal.to_string() }
    } else {
        String::new()
    };

    Ok(EnumItemInfo { name, value })
}

fn get_function_info(type_info: &ITypeInfo, func_desc: &FUNCDESC) -> Result<MethodInfo, Error> {
    let memid = func_desc.memid;
    if memid >= 0x60000000 && memid < 0x60020000 {
        return Err(Error::IoError(std::io::Error::new(
            std::io::ErrorKind::Other,
            "Hidden method",
        )));
    }

    let (name, _) = unsafe { get_type_documentation(type_info, memid) };
    let _invoke_kind = match func_desc.invkind {
        INVOKE_PROPERTYGET => "propget",
        INVOKE_PROPERTYPUT => "propput",
        INVOKE_PROPERTYPUTREF => "propputref",
        _ => "func",
    }
    .to_string();

    let mut ret_type = unsafe { type_desc_to_string(type_info, &func_desc.elemdescFunc.tdesc) };

    let mut names: Vec<BSTR> = vec![BSTR::new(); (func_desc.cParams + 1) as usize];
    let mut c_names = 0;
    unsafe {
        type_info
            .GetNames(memid, names.as_mut_slice(), &mut c_names)
            .ok()
    };

    let mut params = Vec::new();
    for i in 0..func_desc.cParams {
        let elem_desc = unsafe { *func_desc.lprgelemdescParam.offset(i as isize) };
        let param_type = unsafe { type_desc_to_string(type_info, &elem_desc.tdesc) };
        let param_name = if (i + 1) < c_names as i16 {
            names[(i + 1) as usize].to_string()
        } else {
            format!("arg{}", i)
        };

        let param_flags_raw = unsafe { elem_desc.Anonymous.paramdesc.wParamFlags };
        let param_flags = ParamFlags::from_bits_truncate(param_flags_raw.0);
        let mut flags = Vec::new();
        if param_flags.contains(ParamFlags::FIN) {
            flags.push("in".to_string());
        }
        if param_flags.contains(ParamFlags::FOUT) {
            flags.push("out".to_string());
        }
        if param_flags.contains(ParamFlags::FLCID) {
            flags.push("lcid".to_string());
        }
        if param_flags.contains(ParamFlags::FRETVAL) {
            flags.push("retval".to_string());
        }
        if param_flags.contains(ParamFlags::FOPT) {
            flags.push("optional".to_string());
        }
        let mut default_value = None;
        if param_flags.contains(ParamFlags::FHASDEFAULT) {
            flags.push("defaultvalue".to_string());
            let val = unsafe {
                let param_desc_ex = elem_desc.Anonymous.paramdesc.pparamdescex;
                if !param_desc_ex.is_null() {
                    let variant = &(*param_desc_ex).varDefaultValue;
                    variant_to_string(variant)
                } else {
                    String::new()
                }
            };
            if !val.is_empty() {
                default_value = Some(val);
            }
        }

        params.push(ParamInfo {
            name: param_name,
            type_name: param_type,
            flags,
            default_value,
        });
    }

    // Handle return value transformation for HRESULT methods
    if func_desc.invkind == INVOKE_PROPERTYGET || ret_type != "void" {
        let mut real_ret_type = ret_type.clone();
        // Check if there is a retval param
        if let Some(pos) = params
            .iter()
            .position(|p| p.flags.contains(&"retval".to_string()))
        {
            // It's a COM method returning HRESULT with a retval param
            // The "real" return type is the type of the retval param (pointer stripped)
            let retval_param = &params[pos];
            real_ret_type = retval_param.type_name.trim_end_matches('*').to_string();
            // Remove the retval param from the list as it's now the return value
            params.remove(pos);
        }
        ret_type = real_ret_type;
    }

    Ok(MethodInfo {
        name,
        ret_type,
        params,
        _invoke_kind,
    })
}

pub fn get_library_name(tlb_path: &std::path::Path) -> Result<String, Error> {
    let mut type_lib_info = TypeLibInfo::new();
    type_lib_info.load_type_lib(tlb_path)?;
    let (name, _) = type_lib_info.get_documentation(-1)?;
    Ok(name.to_string())
}

pub fn build_tlb<W>(
    tlb_path: &std::path::Path,
    mut out: W,
    import_stdole: bool,
) -> Result<(), Error>
where
    W: std::io::Write,
{
    let mut type_lib_info = TypeLibInfo::new();
    type_lib_info.load_type_lib(tlb_path)?;

    // Use SafeLibAttr which automatically releases on drop
    let lib_attr = type_lib_info.get_lib_attr_safe()?;
    let (name, doc_string) = type_lib_info.get_documentation(-1)?;

    writeln!(out, "// Decompilated from {}", tlb_path.display())?;

    let mut lib_attributes = Vec::new();
    lib_attributes.push(format!("uuid({:?})", lib_attr.guid));
    lib_attributes.push(format!(
        "version({}.{})",
        lib_attr.wMajorVerNum, lib_attr.wMinorVerNum
    ));
    lib_attributes.push(format!("helpstring(\"{}\")", doc_string));

    if let Some(tlib) = &type_lib_info.tlib {
        if let Ok(tlib2) = tlib.cast::<ITypeLib2>() {
            unsafe {
                let custom_attrs = get_lib_custom_data(&tlib2)?;
                lib_attributes.extend(custom_attrs);
            }
        }
    }

    writeln!(out, "[")?;
    for (i, attr) in lib_attributes.iter().enumerate() {
        let suffix = if i == lib_attributes.len() - 1 {
            ""
        } else {
            ","
        };
        writeln!(out, "  {}{}", attr, suffix)?;
    }
    writeln!(out, "]")?;
    writeln!(out, "library {}", name)?;
    writeln!(out, "{{")?;

    // Standard imports often found in IDLs
    if import_stdole {
        writeln!(out, "    importlib(\"stdole2.tlb\");")?;
    }
    writeln!(out, "")?;

    let count = type_lib_info.get_type_info_count();

    // Forward declarations
    for i in 0..count {
        if let Ok(type_info) = type_lib_info.get_type_info(i) {
            if let Ok(type_attr) = SafeTypeAttr::new(&type_info) {
                let type_kind = type_attr.typekind;
                let (name, _) = unsafe { get_type_documentation(&type_info, -1) };

                match type_kind {
                    TKIND_INTERFACE => {
                        writeln!(out, "    interface {};", name)?;
                    }
                    TKIND_DISPATCH => {
                        writeln!(out, "    interface {};", name)?;
                    }
                    TKIND_COCLASS => {
                        writeln!(out, "    coclass {};", name)?;
                    }
                    _ => {}
                }
            }
        }
    }
    writeln!(out, "")?;

    // Enums first
    let count = type_lib_info.get_type_info_count();
    for i in 0..count {
        if let Ok(type_info) = type_lib_info.get_type_info(i) {
            if let Ok(type_attr) = SafeTypeAttr::new(&type_info) {
                if type_attr.typekind == TKIND_ENUM {
                    print_type_info(&type_info, &mut out)?;
                }
            }
        }
    }
    writeln!(out, "")?;

    // Rest of types
    let count = type_lib_info.get_type_info_count();
    for i in 0..count {
        if let Ok(type_info) = type_lib_info.get_type_info(i) {
            if let Ok(type_attr) = SafeTypeAttr::new(&type_info) {
                if type_attr.typekind != TKIND_ENUM {
                    print_type_info(&type_info, &mut out)?;
                }
            }
        }
    }

    writeln!(out, "}};")?;
    Ok(())
}

fn print_interface_header<W>(type_info: &ITypeInfo, out: &mut W) -> Result<(), Error>
where
    W: std::io::Write,
{
    let type_attr = SafeTypeAttr::new(type_info)?;
    let type_kind = type_attr.typekind;
    let guid = type_attr.guid;
    let (_, doc_string) = unsafe { get_type_documentation(type_info, -1) };
    let type_flags = type_attr.wTypeFlags;

    if type_kind == TKIND_INTERFACE
        || type_kind == TKIND_DISPATCH
        || type_kind == TKIND_COCLASS
        || type_kind == TKIND_ENUM
    {
        let mut attributes = Vec::new();
        attributes.push(format!("uuid({:?})", guid));

        if !doc_string.is_empty() {
            attributes.push(format!("helpstring(\"{}\")", doc_string));
        }

        let flags = TypeFlags::from_bits_truncate(type_flags);

        if flags.contains(TypeFlags::FHIDDEN) {
            attributes.push("hidden".to_string());
        }
        if flags.contains(TypeFlags::FDUAL) {
            attributes.push("dual".to_string());
        }
        if flags.contains(TypeFlags::FRESTRICTED) {
            attributes.push("restricted".to_string());
        }
        if flags.contains(TypeFlags::FNONEXTENSIBLE) {
            attributes.push("nonextensible".to_string());
        }
        if flags.contains(TypeFlags::FOLEAUTOMATION) {
            attributes.push("oleautomation".to_string());
        }

        if flags.contains(TypeFlags::FDISPATCHABLE | TypeFlags::FDUAL) {
            attributes.push("oleautomation".to_string());
        }

        // Custom attributes
        if let Ok(type_info2) = type_info.cast::<ITypeInfo2>() {
            unsafe {
                let custom_attrs = get_custom_data(&type_info2)?;
                attributes.extend(custom_attrs);
            }
        }

        writeln!(out, "    [")?;
        for (i, attr) in attributes.iter().enumerate() {
            let suffix = if i == attributes.len() - 1 { "" } else { "," };
            writeln!(out, "      {}{}", attr, suffix)?;
        }
        writeln!(out, "    ]")?;
    }

    Ok(())
}

fn print_type_info<W>(type_info: &ITypeInfo, out: &mut W) -> Result<(), Error>
where
    W: std::io::Write,
{
    let type_attr = SafeTypeAttr::new(type_info)?;

    let guid = type_attr.guid;

    if guid == IUnknown::IID || guid == IDispatch::IID {
        return Ok(());
    }

    let type_kind = type_attr.typekind;
    let type_flags = type_attr.wTypeFlags;

    // Special handling for pure dispinterfaces: extract the inherited interface
    if type_kind == TKIND_DISPATCH
        && !TypeFlags::from_bits_truncate(type_flags).contains(TypeFlags::FDUAL)
    {
        if type_attr.cImplTypes > 0 {
            if let Ok(href) = unsafe { type_info.GetRefTypeOfImplType(0) } {
                if let Ok(ref_type_info) = unsafe { type_info.GetRefTypeInfo(href) } {
                    if let Ok(ref_attr) = SafeTypeAttr::new(&ref_type_info) {
                        let ref_kind = ref_attr.typekind;
                        let ref_flags = ref_attr.wTypeFlags;
                        let is_dual =
                            TypeFlags::from_bits_truncate(ref_flags).contains(TypeFlags::FDUAL);
                        let ref_guid = ref_attr.guid;

                        if (ref_kind == TKIND_INTERFACE || (ref_kind == TKIND_DISPATCH && is_dual))
                            && ref_guid != IUnknown::IID
                            && ref_guid != IDispatch::IID
                        {
                            return print_type_info(&ref_type_info, out);
                        }
                    }
                }
            }
        }
    }

    let (name, _) = unsafe { get_type_documentation(type_info, -1) };

    print_interface_header(type_info, out)?;

    match type_kind {
        TKIND_INTERFACE => print_interface_body(type_info, &type_attr, &name, out)?,
        TKIND_DISPATCH => print_dispatch_body(type_info, &type_attr, &name, out)?,
        TKIND_ENUM => print_enum_body(type_info, &type_attr, &name, out)?,
        TKIND_COCLASS => print_coclass_body(type_info, &type_attr, &name, out)?,
        TKIND_ALIAS => print_alias_body(type_info, &type_attr, &name, out)?,
        TKIND_RECORD => print_record_body(type_info, &type_attr, &name, out)?,
        TKIND_MODULE => print_module_body(type_info, &type_attr, &name, out)?,
        _ => {
            writeln!(out, "    // Unsupported type kind: {:?}", type_kind)?;
        }
    }
    writeln!(out, "")?;

    Ok(())
}

fn print_interface_body<W>(
    type_info: &ITypeInfo,
    type_attr: &TYPEATTR,
    name: &str,
    out: &mut W,
) -> Result<(), Error>
where
    W: std::io::Write,
{
    // Find base interface
    let mut base_name = String::new();
    if type_attr.cImplTypes > 0 {
        if let Ok(href) = unsafe { type_info.GetRefTypeOfImplType(0) } {
            if let Ok(base_info) = unsafe { type_info.GetRefTypeInfo(href) } {
                base_name = unsafe { get_name(&base_info) };
            }
        }
    }

    if !base_name.is_empty() {
        writeln!(out, "    interface {} : {} {{", name, base_name)?;
    } else {
        writeln!(out, "    interface {} {{", name)?;
    }

    // Print properties and methods
    for i in 0..type_attr.cFuncs {
        if let Ok(func_desc) = SafeFuncDesc::new(type_info, i as u32) {
            print_function(type_info, &func_desc, out)?;
        }
    }

    writeln!(out, "    }};")?;
    Ok(())
}

fn print_dispatch_body<W>(
    type_info: &ITypeInfo,
    type_attr: &TYPEATTR,
    name: &str,
    out: &mut W,
) -> Result<(), Error>
where
    W: std::io::Write,
{
    let is_dual = TypeFlags::from_bits_truncate(type_attr.wTypeFlags).contains(TypeFlags::FDUAL);
    if is_dual {
        // Dual interface: get the partner interface (TKIND_INTERFACE)
        // The partner interface is usually at impl type -1 (0xFFFFFFFF)
        let mut partner_type_info = None;
        if let Ok(href) = unsafe { type_info.GetRefTypeOfImplType(u32::MAX) } {
            if let Ok(ref_type_info) = unsafe { type_info.GetRefTypeInfo(href) } {
                partner_type_info = Some(ref_type_info);
            }
        }

        if let Some(partner_info) = partner_type_info {
            // Use the partner interface for everything
            let partner_attr = SafeTypeAttr::new(&partner_info)?;

            // Find base interface of the partner
            let mut base_name = String::new();

            if partner_attr.cImplTypes > 0 {
                if let Ok(href) = unsafe { partner_info.GetRefTypeOfImplType(0) } {
                    if let Ok(base_info) = unsafe { partner_info.GetRefTypeInfo(href) } {
                        base_name = unsafe { get_name(&base_info) };
                    }
                }
            }

            if !base_name.is_empty() {
                writeln!(out, "    interface {} : {} {{", name, base_name)?;
            } else {
                writeln!(out, "    interface {} {{", name)?;
            }

            // Print methods
            for i in 0..partner_attr.cFuncs {
                if let Ok(func_desc) = SafeFuncDesc::new(&partner_info, i as u32) {
                    print_function(&partner_info, &func_desc, out)?;
                }
            }
            writeln!(out, "    }};")?;
        } else {
            // Fallback if partner not found (shouldn't happen for valid duals)
            writeln!(out, "    interface {} : IDispatch {{", name)?;
            for i in 7..type_attr.cFuncs {
                // Skip IDispatch methods
                if let Ok(func_desc) = SafeFuncDesc::new(type_info, i as u32) {
                    print_function(type_info, &func_desc, out)?;
                }
            }
            writeln!(out, "    }};")?;
        }
    } else {
        // Non-dual dispinterface treated as interface : IDispatch
        writeln!(out, "    interface {} : IDispatch {{", name)?;
        for i in 0..type_attr.cFuncs {
            if let Ok(func_desc) = SafeFuncDesc::new(type_info, i as u32) {
                print_function(type_info, &func_desc, out)?;
            }
        }
        writeln!(out, "    }};")?;
    }
    Ok(())
}

fn print_enum_body<W>(
    type_info: &ITypeInfo,
    type_attr: &TYPEATTR,
    name: &str,
    out: &mut W,
) -> Result<(), Error>
where
    W: std::io::Write,
{
    writeln!(out, "    enum {} {{", name)?;
    for i in 0..type_attr.cVars {
        if let Ok(var_desc) = SafeVarDesc::new(type_info, i as u32) {
            print_var(type_info, &var_desc, out)?;
        }
    }
    writeln!(out, "    }};")?;
    Ok(())
}

fn print_coclass_body<W>(
    type_info: &ITypeInfo,
    type_attr: &TYPEATTR,
    name: &str,
    out: &mut W,
) -> Result<(), Error>
where
    W: std::io::Write,
{
    writeln!(out, "    coclass {} {{", name)?;
    for i in 0..type_attr.cImplTypes {
        if let Ok(href) = unsafe { type_info.GetRefTypeOfImplType(i as u32) } {
            if let Ok(ref_type_info) = unsafe { type_info.GetRefTypeInfo(href) } {
                let ref_name = unsafe { get_name(&ref_type_info) };
                let impl_flags_raw = unsafe {
                    type_info
                        .GetImplTypeFlags(i as u32)
                        .unwrap_or(IMPLTYPEFLAGS::default())
                };
                let impl_flags = ImplTypeFlags::from_bits_truncate(impl_flags_raw.0 as u32);
                // Check for [default]
                let default_str = if impl_flags.contains(ImplTypeFlags::FDEFAULT) {
                    "[default] "
                } else {
                    ""
                };
                // Check for [source] (2)
                let source_str = if impl_flags.contains(ImplTypeFlags::FSOURCE) {
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
    Ok(())
}

fn print_alias_body<W>(
    type_info: &ITypeInfo,
    type_attr: &TYPEATTR,
    name: &str,
    out: &mut W,
) -> Result<(), Error>
where
    W: std::io::Write,
{
    let alias_type_name = unsafe { type_desc_to_string(type_info, &type_attr.tdescAlias) };
    let mut attributes = Vec::new();

    if !TypeFlags::from_bits_truncate(type_attr.wTypeFlags).contains(TypeFlags::FHIDDEN) {
        attributes.push("public");
    }

    let attr_str = if !attributes.is_empty() {
        format!("[{}] ", attributes.join(", "))
    } else {
        String::new()
    };

    writeln!(out, "    typedef {}{} {};", attr_str, alias_type_name, name)?;
    Ok(())
}

fn print_record_body<W>(
    type_info: &ITypeInfo,
    type_attr: &TYPEATTR,
    name: &str,
    out: &mut W,
) -> Result<(), Error>
where
    W: std::io::Write,
{
    writeln!(out, "    typedef struct tag{} {{", name)?;
    for i in 0..type_attr.cVars {
        if let Ok(var_desc) = SafeVarDesc::new(type_info, i as u32) {
            print_record_member(type_info, &var_desc, out)?;
        }
    }
    writeln!(out, "    }} {};", name)?;
    Ok(())
}

fn print_module_body<W>(
    type_info: &ITypeInfo,
    type_attr: &TYPEATTR,
    name: &str,
    out: &mut W,
) -> Result<(), Error>
where
    W: std::io::Write,
{
    let mut dll_name = String::new();
    if type_attr.cFuncs > 0 {
        if let Ok(func_desc) = SafeFuncDesc::new(type_info, 0) {
            if let Ok(dll) = unsafe { get_dll_entry(type_info, func_desc.memid, func_desc.invkind) }
            {
                dll_name = dll;
            }
        }
    }

    let mut attributes = Vec::new();
    if !dll_name.is_empty() {
        attributes.push(format!("dllname(\"{}\")", dll_name));
    }

    // UUID and helpstring are already printed in print_interface_header if they exist,
    // but wait, print_interface_header does NOT print for MODULE in the original code?
    // Original code checked for INTERFACE, DISPATCH, COCLASS, ENUM.
    // So for MODULE we need to print attributes here manually as in original.

    attributes.push(format!("uuid({:?})", type_attr.guid));

    let (_, doc_string) = unsafe { get_type_documentation(type_info, -1) };

    if !doc_string.is_empty() {
        attributes.push(format!("helpstring(\"{}\")", doc_string));
    }

    writeln!(out, "    [")?;
    for (i, attr) in attributes.iter().enumerate() {
        let suffix = if i == attributes.len() - 1 { "" } else { "," };
        writeln!(out, "      {}{}", attr, suffix)?;
    }
    writeln!(out, "    ]")?;
    writeln!(out, "    module {} {{", name)?;

    for i in 0..type_attr.cVars {
        if let Ok(var_desc) = SafeVarDesc::new(type_info, i as u32) {
            unsafe {
                print_module_const(type_info, &var_desc, out)?;
            }
        }
    }

    writeln!(out, "    }};")?;
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
        // Should really be variant_to_string but formatted for const
        let val_str = unsafe { variant_to_string(val) };
        // Note: variant_to_string might quote strings, which is good or bad?
        // for const int it should be a number.
        // Original: let val_int = unsafe { val.Anonymous.Anonymous.Anonymous.lVal };
        // writeln!(out, "        const int {} = {};", name, val_int)?;

        // We should try to respect the type in var_desc
        let type_name = unsafe { type_desc_to_string(type_info, &var_desc.elemdescVar.tdesc) };
        writeln!(out, "        const {} {} = {};", type_name, name, val_str)?;
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

unsafe fn get_lib_custom_data(type_lib2: &ITypeLib2) -> Result<Vec<String>, Error> {
    let mut attrs = Vec::new();
    let cust_data = unsafe { type_lib2.GetAllCustData()? };

    for i in 0..cust_data.cCustData {
        let item = unsafe { &*cust_data.prgCustData.offset(i as isize) };
        let guid = item.guid;
        let val = &item.varValue;

        let vt = unsafe { val.Anonymous.Anonymous.vt };
        if vt == VT_BSTR {
            let bstr_val = unsafe { &val.Anonymous.Anonymous.Anonymous.bstrVal };
            let s = bstr_val.to_string();
            attrs.push(format!("custom({:?}, \"{}\")", guid, s));
        }
    }
    Ok(attrs)
}

unsafe fn get_custom_data(type_info2: &ITypeInfo2) -> Result<Vec<String>, Error> {
    let mut attrs = Vec::new();
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
            attrs.push(format!("custom({:?}, \"{}\")", guid, s));
        }
    }
    Ok(attrs)
}

fn print_function<W>(type_info: &ITypeInfo, func_desc: &FUNCDESC, out: &mut W) -> Result<(), Error>
where
    W: std::io::Write,
{
    let memid = func_desc.memid;

    // Filter IUnknown and IDispatch methods
    // Olewoo filters 0x60000000 to 0x60020000
    if memid >= 0x60000000 && memid < 0x60020000 {
        return Ok(());
    }

    let (name, doc_string) = unsafe { get_type_documentation(type_info, memid) };

    let invoke_kind = func_desc.invkind;
    let prop_attr = match invoke_kind {
        INVOKE_PROPERTYGET => "propget",
        INVOKE_PROPERTYPUT => "propput",
        INVOKE_PROPERTYPUTREF => "propputref",
        _ => "",
    };

    let ret_type = unsafe { type_desc_to_string(type_info, &func_desc.elemdescFunc.tdesc) };

    write!(out, "        [id(0x{:08x})", memid)?;
    if !prop_attr.is_empty() {
        write!(out, ", {}", prop_attr)?;
    }
    if !doc_string.is_empty() {
        write!(out, ", helpstring(\"{}\")", doc_string)?;
    }
    write!(out, "]\n")?;

    write!(out, "        {}HRESULT {} (", "", name)?;

    // Get parameter names
    let mut names: Vec<BSTR> = vec![BSTR::new(); (func_desc.cParams + 1) as usize];
    let mut c_names = 0;
    unsafe {
        type_info
            .GetNames(memid, names.as_mut_slice(), &mut c_names)
            .ok();
    }

    let mut has_retval = false;

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
        let param_flags_raw = unsafe { elem_desc.Anonymous.paramdesc.wParamFlags };
        let param_flags = ParamFlags::from_bits_truncate(param_flags_raw.0);
        let mut attrs: Vec<String> = Vec::new();
        if param_flags.contains(ParamFlags::FIN) {
            attrs.push("in".to_string());
        } // PARAMFLAG_FIN
        if param_flags.contains(ParamFlags::FOUT) {
            attrs.push("out".to_string());
        } // PARAMFLAG_FOUT
        if param_flags.contains(ParamFlags::FLCID) {
            attrs.push("lcid".to_string());
        } // PARAMFLAG_FLCID
        if param_flags.contains(ParamFlags::FRETVAL) {
            attrs.push("retval".to_string());
            has_retval = true;
        } // PARAMFLAG_FRETVAL
        if param_flags.contains(ParamFlags::FOPT) {
            attrs.push("optional".to_string());
        } // PARAMFLAG_FOPT
        if param_flags.contains(ParamFlags::FHASDEFAULT) {
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

        // Ensure [in] is present if optional is set and no direction is specified
        if attrs.iter().any(|a| a == "optional") && !attrs.iter().any(|a| a == "in" || a == "out") {
            attrs.insert(0, "in".to_string());
        }

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

    if (invoke_kind == INVOKE_PROPERTYGET || ret_type != "void") && !has_retval {
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

fn print_var<W>(type_info: &ITypeInfo, var_desc: &VARDESC, out: &mut W) -> Result<(), Error>
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

fn print_record_member<W>(
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
            if unsafe { tdesc.Anonymous.lptdesc.is_null() } {
                "void*".to_string()
            } else {
                let pointed_type =
                    unsafe { type_desc_to_string(type_info, &*tdesc.Anonymous.lptdesc) };
                format!("{}*", pointed_type)
            }
        }
        VT_SAFEARRAY => {
            if unsafe { tdesc.Anonymous.lptdesc.is_null() } {
                "SAFEARRAY(void)".to_string()
            } else {
                let element_type =
                    unsafe { type_desc_to_string(type_info, &*tdesc.Anonymous.lptdesc) };
                format!("SAFEARRAY({})", element_type)
            }
        }
        VT_USERDEFINED => {
            if let Ok(ref_type_info) = unsafe { type_info.GetRefTypeInfo(tdesc.Anonymous.hreftype) }
            {
                let name = unsafe { get_name(&ref_type_info) };
                if let Ok(ref_attr) = SafeTypeAttr::new(&ref_type_info) {
                    if ref_attr.typekind == TKIND_ENUM {
                        format!("enum {}", name)
                    } else {
                        name
                    }
                } else {
                    name
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
            VT_I1 => variant.Anonymous.Anonymous.Anonymous.cVal.to_string(),
            VT_UI1 => variant.Anonymous.Anonymous.Anonymous.bVal.to_string(),
            VT_UI2 => variant.Anonymous.Anonymous.Anonymous.uiVal.to_string(),
            VT_UI4 => variant.Anonymous.Anonymous.Anonymous.ulVal.to_string(),
            VT_INT => variant.Anonymous.Anonymous.Anonymous.intVal.to_string(),
            VT_UINT => variant.Anonymous.Anonymous.Anonymous.uintVal.to_string(),
            VT_CY => {
                let cy_val = variant.Anonymous.Anonymous.Anonymous.cyVal.int64;
                // Currency is int64 scaled by 10000
                format!("{}", cy_val as f64 / 10000.0)
            }
            VT_DATE => variant.Anonymous.Anonymous.Anonymous.date.to_string(),
            VT_ERROR => format!("0x{:X}", variant.Anonymous.Anonymous.Anonymous.scode),
            _ => format!("/* vt: {} */", variant.Anonymous.Anonymous.vt.0),
        }
    }
}
