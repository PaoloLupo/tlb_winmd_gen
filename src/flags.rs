use bitflags::bitflags;

bitflags! {
    #[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
    pub struct TypeFlags: u16 {
        const FAPPOBJECT = 0x01;
        const FCANCREATE = 0x02;
        const FLICENSED = 0x04;
        const FPREDECLID = 0x08;
        const FHIDDEN = 0x10;
        const FCONTROL = 0x20;
        const FDUAL = 0x40;
        const FNONEXTENSIBLE = 0x80;
        const FOLEAUTOMATION = 0x100;
        const FRESTRICTED = 0x200;
        const FAGGREGATABLE = 0x400;
        const FREPLACEABLE = 0x800;
        const FDISPATCHABLE = 0x1000;
        const FREVERSEBIND = 0x2000;
        const FPROXY = 0x4000;
    }
}

bitflags! {
    #[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
    pub struct ParamFlags: u16 {
        const FIN = 0x01;
        const FOUT = 0x02;
        const FLCID = 0x04;
        const FRETVAL = 0x08;
        const FOPT = 0x10;
        const FHASDEFAULT = 0x20;
        const FHASCUSTDATA = 0x40;
    }
}

bitflags! {
    #[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
    pub struct ImplTypeFlags: u32 {
        const FDEFAULT = 0x01;
        const FSOURCE = 0x02;
        const FRESTRICTED = 0x04;
        const FDEFAULTVTABLE = 0x08;
    }
}
