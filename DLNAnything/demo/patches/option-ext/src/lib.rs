//! Derived from VibeOffice desktop/patches/option-ext (Apache-2.0).
//! The single extension used by dirs-sys.
pub trait OptionExt<T> {
    fn contains<U: PartialEq<T>>(&self, other: &U) -> bool;
}
impl<T> OptionExt<T> for Option<T> {
    fn contains<U: PartialEq<T>>(&self, other: &U) -> bool {
        self.as_ref().is_some_and(|value| other.eq(value))
    }
}
