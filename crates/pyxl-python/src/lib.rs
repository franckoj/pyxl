#![allow(clippy::useless_conversion)]

use pyo3::prelude::*;

mod bindings;

pub use bindings::{PyCell, PyWorkbook, PyWorksheet};

#[pymodule]
fn _core(_py: Python<'_>, m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<PyWorkbook>()?;
    m.add_class::<PyWorksheet>()?;
    m.add_class::<PyCell>()?;
    Ok(())
}
