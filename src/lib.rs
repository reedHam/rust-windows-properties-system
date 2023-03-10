#![feature(test)]
use std::io::Error;
use std::path::Path;

use windows::core::{HSTRING, PCWSTR, PWSTR};
use windows::Win32::System::Com::StructuredStorage::{PropVariantClear, PROPVARIANT};
use windows::Win32::System::Com::*;
use windows::Win32::UI::Shell::PropertiesSystem::*;
use windows::*;

const DEFAULT_PROP_STRING: PCWSTR = w!("");
const DEFAULT_PROP_U32: u32 = 0;

pub trait FromPropVariant {
    fn from_prop_variant(prop: PROPVARIANT) -> Self;
}

impl FromPropVariant for String {
    fn from_prop_variant(mut prop: PROPVARIANT) -> Self {
        unsafe {
            let prop_sting: PWSTR = PropVariantToStringWithDefault(&prop, DEFAULT_PROP_STRING);
            PropVariantClear(&mut prop).unwrap();
            prop_sting.to_string().unwrap_or("".to_string())
        }
    }
}

impl FromPropVariant for u32 {
    fn from_prop_variant(mut prop: PROPVARIANT) -> Self {
        unsafe {
            let prop_u32 = PropVariantToUInt32WithDefault(&prop, DEFAULT_PROP_U32);
            PropVariantClear(&mut prop).unwrap();
            prop_u32
        }
    }
}

pub trait ToPropVariant {
    fn to_prop_variant(&self) -> PROPVARIANT;
}

impl ToPropVariant for String {
    fn to_prop_variant(&self) -> PROPVARIANT {
        unsafe {
            let string = PCWSTR::from_raw(HSTRING::from(self).as_wide().as_ptr());
            InitPropVariantFromStringAsVector(string).unwrap()
        }
    }
}

impl ToPropVariant for &str {
    fn to_prop_variant(&self) -> PROPVARIANT {
        unsafe {
            let string = PCWSTR::from_raw(HSTRING::from(self.to_string()).as_wide().as_ptr());
            InitPropVariantFromStringAsVector(string).unwrap()
        }
    }
}

impl ToPropVariant for Vec<&str> {
    fn to_prop_variant(&self) -> PROPVARIANT {
        unsafe {
            let mut str_vec: Vec<PCWSTR> = Vec::new();

            for element in self {
                str_vec.push(PCWSTR::from_raw(HSTRING::from(*element).as_wide().as_ptr()));
            }

            let str_array = str_vec.as_slice();

            InitPropVariantFromStringVector(Some(str_array)).unwrap()
        }
    }
}

pub struct PropVector {
    pub vector: Vec<String>,
}

impl FromPropVariant for PropVector {
    fn from_prop_variant(prop: PROPVARIANT) -> Self {
        let element_count = &mut 0;
        unsafe {
            let val_vec: *mut *mut PWSTR = &mut std::ptr::null_mut();
            let result = PropVariantToStringVectorAlloc(&prop, val_vec, element_count);
            if result.is_err() {
                return Self { vector: Vec::new() };
            }
            let ptr_vector =
                Vec::from_raw_parts(*val_vec, *element_count as usize, *element_count as usize);
            let string_vector = ptr_vector
                .iter()
                .map(|x| {
                    let string = x.to_string().unwrap();
                    CoTaskMemFree(Some(x.as_ptr() as *const std::ffi::c_void));
                    string
                })
                .collect();
            Self {
                vector: string_vector,
            }
        }
    }
}

impl std::fmt::Display for PropVector {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let mut string = String::new();
        self.vector.iter().for_each(|x| {
            string.push_str(x);
            string.push_str(", ");
        });
        write!(f, "{}", string)
    }
}

pub struct FileProperties {
    path: HSTRING,
    props: IPropertyStore,
    context: IBindCtx,
}

impl FileProperties {
    pub fn new(path: &str, flag: Option<GETPROPERTYSTOREFLAGS>) -> Result<Self, Error> {
        if !Path::exists(Path::new(path)) {
            return Err(Error::new(
                std::io::ErrorKind::NotFound,
                format!("{} not found", path),
            ));
        }
        let path: HSTRING = HSTRING::from(path);
        unsafe {
            CoInitializeEx(None, COINIT_APARTMENTTHREADED)?;
            let context = CreateBindCtx(0)?;

            let flag = flag.unwrap_or(GPS_READWRITE);

            let props: IPropertyStore = SHGetPropertyStoreFromParsingName(&path, &context, flag)?;

            Ok(Self {
                path,
                props,
                context,
            })
        }
    }

    pub fn get_prop_count(&self) -> u32 {
        unsafe { self.props.GetCount().unwrap() }
    }

    pub fn get_prop<T: FromPropVariant>(&self, prop_name: &str) -> Result<T, Error> {
        let prop_key = &mut PROPERTYKEY::default();
        let prop_name = HSTRING::from(prop_name);
        unsafe {
            PSGetPropertyKeyFromName(&prop_name, prop_key)?;
            let prop_variant = self.props.GetValue(prop_key)?;
            Ok(T::from_prop_variant(prop_variant))
        }
    }

    pub fn set_prop<T: ToPropVariant>(&self, prop_name: &str, value: T) -> Result<(), Error> {
        let prop_name = HSTRING::from(prop_name);
        let tag_prop_key = &mut PROPERTYKEY::default();
        unsafe {
            PSGetPropertyKeyFromName(&prop_name, tag_prop_key)?;
            let mut prop_var = value.to_prop_variant();
            PSCoerceToCanonicalValue(tag_prop_key, &mut prop_var)?;
            self.props.SetValue(tag_prop_key, &prop_var)?;
            Ok(())
        }
    }

    pub fn commit(&self) -> Result<(), Error> {
        unsafe { self.props.Commit()? };
        Ok(())
    }
}

impl Drop for FileProperties {
    fn drop(&mut self) {
        unsafe {
            self.context.ReleaseBoundObjects().unwrap();
            CoUninitialize();
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::fs;

    extern crate test;
    use test::Bencher;

    const TEST_FILE_DIR: &str = r#".\test"#;

    fn get_full_path(file_name: &str) -> String {
        Path::new(file_name)
            .canonicalize()
            .unwrap()
            .into_os_string()
            .into_string()
            .unwrap()
            .replace(r#"\\?\"#, "")
    }

    fn enumerate_test_files() -> impl Iterator<Item = String> {
        fs::read_dir(get_full_path(TEST_FILE_DIR))
            .unwrap()
            .map(|x| x.unwrap().path().to_str().unwrap().to_string())
            .filter(|x| x.ends_with(".mp4"))
    }

    #[test]
    fn gets_props() {
        for file in enumerate_test_files() {
            let props = FileProperties::new(&file, None).unwrap();
            let id: String = props.get_prop("System.Media.UniqueFileIdentifier").unwrap();
            if file.contains("without") {
                assert!(id.is_empty());
            } else {
                assert!(!id.is_empty());
            }
        }
    }

    #[test]
    fn sets_props() {
        let raw_test_video_path = format!("{}\\{}", TEST_FILE_DIR, "video_without_properties.mp4");
        let raw_test_video_path = Path::new(&raw_test_video_path).canonicalize().unwrap();
        let full_test_dir_path = raw_test_video_path.parent().unwrap();
        let full_test_file_path = Path::join(full_test_dir_path, "new_test_video.mp4")
            .into_os_string()
            .into_string()
            .unwrap()
            .replace(r#"\\?\"#, "");

        fs::copy(raw_test_video_path, &full_test_file_path).unwrap();

        let test_id = "this_is_the_test_id";

        {
            let props = FileProperties::new(&full_test_file_path, Some(GPS_READWRITE)).unwrap();
            props
                .set_prop("System.Media.UniqueFileIdentifier", test_id)
                .unwrap();
            props.commit().unwrap();
        }

        {
            let props = FileProperties::new(&full_test_file_path, None).unwrap();
            let id: String = props.get_prop("System.Media.UniqueFileIdentifier").unwrap();

            assert_eq!(id, test_id);
        }

        fs::remove_file(&full_test_file_path).unwrap();
    }

    #[bench]
    fn bench_get_string_prop(b: &mut Bencher) {
        let files = enumerate_test_files().collect::<Vec<_>>();
        b.iter(|| {
            for file in &files {
                let props = FileProperties::new(file, None).unwrap();
                let id = props
                    .get_prop::<String>("System.Media.UniqueFileIdentifier")
                    .unwrap();

                if file.contains("without") {
                    assert!(id.is_empty());
                } else {
                    assert!(!id.is_empty());
                }
            }
        })
    }

    #[bench]
    fn bench_get_u32_prop(b: &mut Bencher) {
        let files = enumerate_test_files().collect::<Vec<_>>();
        b.iter(|| {
            for file in &files {
                let props = FileProperties::new(file, None).unwrap();
                props.get_prop::<u32>("System.Generic.Integer").unwrap();
            }
        })
    }

    #[bench]
    fn bench_get_vector_prop(b: &mut Bencher) {
        let files = enumerate_test_files().collect::<Vec<_>>();
        b.iter(|| {
            for file in &files {
                let props = FileProperties::new(file, None).unwrap();
                props.get_prop::<PropVector>("System.Author").unwrap();
            }
        })
    }
}
