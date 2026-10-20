use std::collections::HashMap;
use std::sync::Arc;

use parking_lot::RwLock;
use pyo3::exceptions::{
    PyIndexError, PyKeyError, PyNotImplementedError, PyTypeError, PyValueError,
};
use pyo3::prelude::*;
use pyo3::types::{PyBool, PyFloat, PyInt, PyList, PyString, PyTuple};

use pyxl_core::cell::CellValue;
use pyxl_core::coordinate::{
    coordinate_to_tuple, get_column_letter, range_boundaries, tuple_to_coordinate,
};
use pyxl_core::error::{PyxlError, PyxlResult};
use pyxl_core::workbook::Workbook;
use pyxl_ooxml::{
    read_single_cell, SharedStringsTable, SheetInfo, StreamingRowReader, Styles, XlsxPackage,
};

pub fn map_pyxl_err(err: PyxlError) -> PyErr {
    match err {
        PyxlError::SheetNotFound(name) => {
            PyKeyError::new_err(format!("Worksheet {} does not exist.", name))
        }
        PyxlError::DuplicateSheetTitle(name) => {
            PyValueError::new_err(format!("Worksheet {} already exists.", name))
        }
        PyxlError::InvalidSheetTitle(msg) => PyValueError::new_err(msg),
        PyxlError::SheetIndexOutOfBounds(idx, count) => PyIndexError::new_err(format!(
            "Worksheet index {} out of range (sheet count {})",
            idx, count
        )),
        PyxlError::CannotRemoveLastSheet => {
            PyValueError::new_err("Cannot remove the only worksheet from workbook")
        }
        PyxlError::InvalidCoordinate(coord) => {
            PyValueError::new_err(format!("Invalid coordinate '{}'", coord))
        }
        PyxlError::InvalidRange(rng) => PyValueError::new_err(format!("Invalid range '{}'", rng)),
        PyxlError::RowOutOfBounds(row) => PyValueError::new_err(format!(
            "Row {} is out of bounds (allowed: 1-1,048,576)",
            row
        )),
        PyxlError::ColumnOutOfBounds(col) => PyValueError::new_err(format!(
            "Column {} is out of bounds (allowed: 1-16,384)",
            col
        )),
        PyxlError::ValueError(msg) => PyValueError::new_err(msg),
        PyxlError::TypeError(msg) => PyTypeError::new_err(msg),
        PyxlError::IoError(msg) => pyo3::exceptions::PyIOError::new_err(msg),
    }
}

pub fn python_value_to_cell_value(val: &Bound<'_, PyAny>) -> PyResult<CellValue> {
    if val.is_none() {
        Ok(CellValue::Empty)
    } else if val.is_instance_of::<PyBool>() {
        Ok(CellValue::Bool(val.extract::<bool>()?))
    } else if val.is_instance_of::<PyInt>() {
        Ok(CellValue::Int(val.extract::<i64>()?))
    } else if val.is_instance_of::<PyFloat>() {
        Ok(CellValue::Float(val.extract::<f64>()?))
    } else if val.is_instance_of::<PyString>() {
        let s = val.extract::<String>()?;
        if s.starts_with('=') {
            Ok(CellValue::Formula {
                formula: s,
                cached_value: None,
            })
        } else {
            Ok(CellValue::String(s))
        }
    } else if val.hasattr("strftime")? {
        let s = val.str()?.to_string();
        if val.hasattr("hour")? {
            Ok(CellValue::DateTime(s))
        } else {
            Ok(CellValue::Date(s))
        }
    } else {
        // Fallback: convert to string representation
        let s = val.str()?.to_string();
        Ok(CellValue::String(s))
    }
}

pub fn cell_value_to_python<'py>(py: Python<'py>, val: &CellValue) -> PyResult<Bound<'py, PyAny>> {
    match val {
        CellValue::Empty => Ok(py.None().into_bound(py)),
        CellValue::String(s) => Ok(s.to_object(py).into_bound(py)),
        CellValue::Int(i) => Ok(i.to_object(py).into_bound(py)),
        CellValue::Float(f) => Ok(f.to_object(py).into_bound(py)),
        CellValue::Bool(b) => Ok(b.to_object(py).into_bound(py)),
        CellValue::Date(d) => date_str_to_pydatetime(py, d),
        CellValue::DateTime(dt) => date_str_to_pydatetime(py, dt),
        CellValue::Formula { formula, .. } => Ok(formula.to_object(py).into_bound(py)),
    }
}

fn date_str_to_pydatetime<'py>(py: Python<'py>, s: &str) -> PyResult<Bound<'py, PyAny>> {
    if s.len() >= 10 && s.as_bytes()[4] == b'-' && s.as_bytes()[7] == b'-' {
        let year: i32 = s[0..4].parse().unwrap_or(1970);
        let month: u8 = s[5..7].parse().unwrap_or(1);
        let day: u8 = s[8..10].parse().unwrap_or(1);
        let (hour, min, sec, micro) =
            if s.len() >= 19 && (s.as_bytes()[10] == b' ' || s.as_bytes()[10] == b'T') {
                let h: u8 = s[11..13].parse().unwrap_or(0);
                let m: u8 = s[14..16].parse().unwrap_or(0);
                let sec_str = &s[17..];
                if let Some((s_part, us_part)) = sec_str.split_once('.') {
                    let s_val: u8 = s_part.parse().unwrap_or(0);
                    let us_val: u32 = us_part.parse().unwrap_or(0);
                    (h, m, s_val, us_val)
                } else {
                    let s_val: u8 = sec_str.get(..2).and_then(|x| x.parse().ok()).unwrap_or(0);
                    (h, m, s_val, 0)
                }
            } else {
                (0, 0, 0, 0)
            };
        let dt_mod = py.import_bound("datetime")?;
        let dt_cls = dt_mod.getattr("datetime")?;
        if let Ok(dt) = dt_cls.call1((year, month, day, hour, min, sec, micro)) {
            return Ok(dt);
        }
    }
    Ok(s.to_object(py).into_bound(py))
}

#[pyclass(name = "Cell")]
pub struct PyCell {
    wb: Arc<RwLock<Workbook>>,
    sheet_title: Arc<RwLock<String>>,
    row: u32,
    column: u32,
}

#[pymethods]
impl PyCell {
    #[getter]
    pub fn row(&self) -> u32 {
        self.row
    }

    #[getter]
    pub fn column(&self) -> u32 {
        self.column
    }

    #[getter]
    pub fn coordinate(&self) -> PyResult<String> {
        tuple_to_coordinate(self.row, self.column).map_err(map_pyxl_err)
    }

    #[getter]
    pub fn column_letter(&self) -> PyResult<String> {
        get_column_letter(self.column).map_err(map_pyxl_err)
    }

    #[getter]
    pub fn data_type(&self) -> PyResult<&'static str> {
        let wb = self.wb.read();
        let title = self.sheet_title.read().clone();
        let ws = wb
            .get_sheet(&title)
            .ok_or_else(|| PyxlError::SheetNotFound(title.clone()))
            .map_err(map_pyxl_err)?;
        let val = ws.get_cell_value(self.row, self.column);
        Ok(val.data_type())
    }

    #[getter]
    pub fn value<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyAny>> {
        let wb = self.wb.read();
        let data_only = wb.is_data_only();
        let title = self.sheet_title.read().clone();
        let ws = wb
            .get_sheet(&title)
            .ok_or_else(|| PyxlError::SheetNotFound(title.clone()))
            .map_err(map_pyxl_err)?;
        let val = ws.get_resolved_cell_value(self.row, self.column, data_only);
        cell_value_to_python(py, &val)
    }

    #[setter]
    pub fn set_value(&self, new_val: &Bound<'_, PyAny>) -> PyResult<()> {
        let val = python_value_to_cell_value(new_val)?;
        let mut wb = self.wb.write();
        let title = self.sheet_title.read().clone();
        let ws = wb
            .get_sheet_mut(&title)
            .ok_or_else(|| PyxlError::SheetNotFound(title.clone()))
            .map_err(map_pyxl_err)?;
        ws.set_cell_value(self.row, self.column, val)
            .map_err(map_pyxl_err)?;
        Ok(())
    }

    pub fn __repr__(&self) -> PyResult<String> {
        let title = self.sheet_title.read().clone();
        let coord = tuple_to_coordinate(self.row, self.column).map_err(map_pyxl_err)?;
        Ok(format!("<Cell '{}'.{}>", title, coord))
    }
}

#[pyclass(name = "Worksheet")]
pub struct PyWorksheet {
    wb: Arc<RwLock<Workbook>>,
    sheet_title: Arc<RwLock<String>>,
}

#[pymethods]
impl PyWorksheet {
    #[getter]
    pub fn title(&self) -> String {
        self.sheet_title.read().clone()
    }

    #[setter]
    pub fn set_title(&self, new_title: &str) -> PyResult<()> {
        let old_title = self.sheet_title.read().clone();
        let mut wb = self.wb.write();
        wb.rename_sheet(&old_title, new_title)
            .map_err(map_pyxl_err)?;
        *self.sheet_title.write() = new_title.to_string();
        Ok(())
    }

    #[getter]
    pub fn min_row(&self) -> PyResult<u32> {
        let wb = self.wb.read();
        let title = self.sheet_title.read().clone();
        let ws = wb
            .get_sheet(&title)
            .ok_or_else(|| PyxlError::SheetNotFound(title.clone()))
            .map_err(map_pyxl_err)?;
        Ok(ws.min_row())
    }

    #[getter]
    pub fn max_row(&self) -> PyResult<u32> {
        let wb = self.wb.read();
        let title = self.sheet_title.read().clone();
        let ws = wb
            .get_sheet(&title)
            .ok_or_else(|| PyxlError::SheetNotFound(title.clone()))
            .map_err(map_pyxl_err)?;
        Ok(ws.max_row())
    }

    #[getter]
    pub fn min_column(&self) -> PyResult<u32> {
        let wb = self.wb.read();
        let title = self.sheet_title.read().clone();
        let ws = wb
            .get_sheet(&title)
            .ok_or_else(|| PyxlError::SheetNotFound(title.clone()))
            .map_err(map_pyxl_err)?;
        Ok(ws.min_column())
    }

    #[getter]
    pub fn max_column(&self) -> PyResult<u32> {
        let wb = self.wb.read();
        let title = self.sheet_title.read().clone();
        let ws = wb
            .get_sheet(&title)
            .ok_or_else(|| PyxlError::SheetNotFound(title.clone()))
            .map_err(map_pyxl_err)?;
        Ok(ws.max_column())
    }

    #[getter]
    pub fn dimensions(&self) -> PyResult<String> {
        let wb = self.wb.read();
        let title = self.sheet_title.read().clone();
        let ws = wb
            .get_sheet(&title)
            .ok_or_else(|| PyxlError::SheetNotFound(title.clone()))
            .map_err(map_pyxl_err)?;
        ws.dimensions().map_err(map_pyxl_err)
    }

    #[pyo3(signature = (row, column, value=None))]
    pub fn cell(
        &self,
        row: u32,
        column: u32,
        value: Option<&Bound<'_, PyAny>>,
    ) -> PyResult<PyCell> {
        {
            let mut wb = self.wb.write();
            let title = self.sheet_title.read().clone();
            let ws = wb
                .get_sheet_mut(&title)
                .ok_or_else(|| PyxlError::SheetNotFound(title.clone()))
                .map_err(map_pyxl_err)?;

            if let Some(v) = value {
                let cell_val = python_value_to_cell_value(v)?;
                ws.set_cell_value(row, column, cell_val)
                    .map_err(map_pyxl_err)?;
            } else {
                ws.get_or_create_cell(row, column).map_err(map_pyxl_err)?;
            }
        }

        Ok(PyCell {
            wb: Arc::clone(&self.wb),
            sheet_title: Arc::clone(&self.sheet_title),
            row,
            column,
        })
    }

    pub fn __getitem__<'py>(
        &self,
        py: Python<'py>,
        key: &Bound<'py, PyAny>,
    ) -> PyResult<Bound<'py, PyAny>> {
        let key_str = key.extract::<String>()?;
        if key_str.contains(':') {
            let (min_r, min_c, max_r, max_c) = range_boundaries(&key_str).map_err(map_pyxl_err)?;
            let mut outer_rows = Vec::with_capacity((max_r - min_r + 1) as usize);

            for r in min_r..=max_r {
                let mut inner_cols = Vec::with_capacity((max_c - min_c + 1) as usize);
                for c in min_c..=max_c {
                    let cell = Bound::new(
                        py,
                        PyCell {
                            wb: Arc::clone(&self.wb),
                            sheet_title: Arc::clone(&self.sheet_title),
                            row: r,
                            column: c,
                        },
                    )?;
                    inner_cols.push(cell.into_any());
                }
                outer_rows.push(PyTuple::new_bound(py, inner_cols));
            }
            Ok(PyTuple::new_bound(py, outer_rows).into_any())
        } else {
            let (r, c) = coordinate_to_tuple(&key_str).map_err(map_pyxl_err)?;
            {
                let mut wb = self.wb.write();
                let title = self.sheet_title.read().clone();
                let ws = wb
                    .get_sheet_mut(&title)
                    .ok_or_else(|| PyxlError::SheetNotFound(title.clone()))
                    .map_err(map_pyxl_err)?;
                ws.get_or_create_cell(r, c).map_err(map_pyxl_err)?;
            }
            let cell = Bound::new(
                py,
                PyCell {
                    wb: Arc::clone(&self.wb),
                    sheet_title: Arc::clone(&self.sheet_title),
                    row: r,
                    column: c,
                },
            )?;
            Ok(cell.into_any())
        }
    }

    pub fn __setitem__(&self, key: &str, value: &Bound<'_, PyAny>) -> PyResult<()> {
        let (r, c) = coordinate_to_tuple(key).map_err(map_pyxl_err)?;
        let cell_val = python_value_to_cell_value(value)?;
        let mut wb = self.wb.write();
        let title = self.sheet_title.read().clone();
        let ws = wb
            .get_sheet_mut(&title)
            .ok_or_else(|| PyxlError::SheetNotFound(title.clone()))
            .map_err(map_pyxl_err)?;
        ws.set_cell_value(r, c, cell_val).map_err(map_pyxl_err)?;
        Ok(())
    }

    pub fn append(&self, iterable: &Bound<'_, PyAny>) -> PyResult<u32> {
        let mut values = Vec::new();
        for item in iterable.iter()? {
            let item_val = item?;
            values.push(python_value_to_cell_value(&item_val)?);
        }

        let mut wb = self.wb.write();
        let title = self.sheet_title.read().clone();
        let ws = wb
            .get_sheet_mut(&title)
            .ok_or_else(|| PyxlError::SheetNotFound(title.clone()))
            .map_err(map_pyxl_err)?;
        ws.append(values).map_err(map_pyxl_err)
    }

    #[pyo3(signature = (min_row=None, max_row=None, min_col=None, max_col=None, values_only=false))]
    pub fn iter_rows<'py>(
        &self,
        py: Python<'py>,
        min_row: Option<u32>,
        max_row: Option<u32>,
        min_col: Option<u32>,
        max_col: Option<u32>,
        values_only: bool,
    ) -> PyResult<Bound<'py, PyAny>> {
        let (r_start, r_end, c_start, c_end) = {
            let wb = self.wb.read();
            let title = self.sheet_title.read().clone();
            let ws = wb
                .get_sheet(&title)
                .ok_or_else(|| PyxlError::SheetNotFound(title.clone()))
                .map_err(map_pyxl_err)?;
            ws.resolve_bounds(min_row, max_row, min_col, max_col)
        };

        if r_start > r_end || c_start > c_end {
            return Ok(PyList::empty_bound(py).into_any());
        }

        let mut rows = Vec::with_capacity((r_end - r_start + 1) as usize);

        if values_only {
            let wb = self.wb.read();
            let data_only = wb.is_data_only();
            let title = self.sheet_title.read().clone();
            let ws = wb
                .get_sheet(&title)
                .ok_or_else(|| PyxlError::SheetNotFound(title.clone()))
                .map_err(map_pyxl_err)?;

            for r in r_start..=r_end {
                let mut row_vals = Vec::with_capacity((c_end - c_start + 1) as usize);
                for c in c_start..=c_end {
                    let val = ws.get_resolved_cell_value(r, c, data_only);
                    row_vals.push(cell_value_to_python(py, &val)?);
                }
                rows.push(PyTuple::new_bound(py, row_vals));
            }
        } else {
            for r in r_start..=r_end {
                let mut row_cells = Vec::with_capacity((c_end - c_start + 1) as usize);
                for c in c_start..=c_end {
                    let cell = Bound::new(
                        py,
                        PyCell {
                            wb: Arc::clone(&self.wb),
                            sheet_title: Arc::clone(&self.sheet_title),
                            row: r,
                            column: c,
                        },
                    )?;
                    row_cells.push(cell.into_any());
                }
                rows.push(PyTuple::new_bound(py, row_cells));
            }
        }

        let py_list = PyList::new_bound(py, rows);
        Ok(py_list.into_any())
    }

    #[pyo3(signature = (min_row=None, max_row=None, min_col=None, max_col=None, values_only=false))]
    pub fn iter_cols<'py>(
        &self,
        py: Python<'py>,
        min_row: Option<u32>,
        max_row: Option<u32>,
        min_col: Option<u32>,
        max_col: Option<u32>,
        values_only: bool,
    ) -> PyResult<Bound<'py, PyAny>> {
        let (r_start, r_end, c_start, c_end) = {
            let wb = self.wb.read();
            let title = self.sheet_title.read().clone();
            let ws = wb
                .get_sheet(&title)
                .ok_or_else(|| PyxlError::SheetNotFound(title.clone()))
                .map_err(map_pyxl_err)?;
            ws.resolve_bounds(min_row, max_row, min_col, max_col)
        };

        if r_start > r_end || c_start > c_end {
            return Ok(PyList::empty_bound(py).into_any());
        }

        let mut cols = Vec::with_capacity((c_end - c_start + 1) as usize);

        if values_only {
            let wb = self.wb.read();
            let data_only = wb.is_data_only();
            let title = self.sheet_title.read().clone();
            let ws = wb
                .get_sheet(&title)
                .ok_or_else(|| PyxlError::SheetNotFound(title.clone()))
                .map_err(map_pyxl_err)?;

            for c in c_start..=c_end {
                let mut col_vals = Vec::with_capacity((r_end - r_start + 1) as usize);
                for r in r_start..=r_end {
                    let val = ws.get_resolved_cell_value(r, c, data_only);
                    col_vals.push(cell_value_to_python(py, &val)?);
                }
                cols.push(PyTuple::new_bound(py, col_vals));
            }
        } else {
            for c in c_start..=c_end {
                let mut col_cells = Vec::with_capacity((r_end - r_start + 1) as usize);
                for r in r_start..=r_end {
                    let cell = Bound::new(
                        py,
                        PyCell {
                            wb: Arc::clone(&self.wb),
                            sheet_title: Arc::clone(&self.sheet_title),
                            row: r,
                            column: c,
                        },
                    )?;
                    col_cells.push(cell.into_any());
                }
                cols.push(PyTuple::new_bound(py, col_cells));
            }
        }

        let py_list = PyList::new_bound(py, cols);
        Ok(py_list.into_any())
    }

    #[getter]
    pub fn values<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyAny>> {
        self.iter_rows(py, None, None, None, None, true)
    }

    #[getter]
    pub fn rows<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyAny>> {
        self.iter_rows(py, None, None, None, None, false)
    }

    #[getter]
    pub fn columns<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyAny>> {
        self.iter_cols(py, None, None, None, None, false)
    }

    pub fn __repr__(&self) -> String {
        format!("<Worksheet \"{}\">", self.sheet_title.read())
    }
}

#[pyclass(name = "Workbook")]
pub struct PyWorkbook {
    wb: Arc<RwLock<Workbook>>,
}

#[pymethods]
impl PyWorkbook {
    #[new]
    pub fn new() -> Self {
        Self {
            wb: Arc::new(RwLock::new(Workbook::new())),
        }
    }

    #[getter]
    pub fn read_only(&self) -> bool {
        self.wb.read().is_read_only()
    }

    #[getter]
    pub fn data_only(&self) -> bool {
        self.wb.read().is_data_only()
    }

    #[getter]
    pub fn active(&self) -> PyResult<PyWorksheet> {
        let wb = self.wb.read();
        let title = wb.active().title.clone();
        Ok(PyWorksheet {
            wb: Arc::clone(&self.wb),
            sheet_title: Arc::new(RwLock::new(title)),
        })
    }

    #[setter]
    pub fn set_active(&self, target: &Bound<'_, PyAny>) -> PyResult<()> {
        let mut wb = self.wb.write();
        if let Ok(idx) = target.extract::<usize>() {
            wb.set_active_index(idx).map_err(map_pyxl_err)?;
        } else if let Ok(ws) = target.extract::<PyRef<'_, PyWorksheet>>() {
            let title = ws.sheet_title.read().clone();
            wb.set_active_by_title(&title).map_err(map_pyxl_err)?;
        } else if let Ok(title) = target.extract::<String>() {
            wb.set_active_by_title(&title).map_err(map_pyxl_err)?;
        } else {
            return Err(PyTypeError::new_err(
                "active sheet must be set with an index, sheet name, or Worksheet object",
            ));
        }
        Ok(())
    }

    #[getter]
    pub fn sheetnames(&self) -> Vec<String> {
        self.wb.read().sheetnames()
    }

    #[getter]
    pub fn worksheets(&self) -> Vec<PyWorksheet> {
        let wb = self.wb.read();
        wb.sheetnames()
            .into_iter()
            .map(|title| PyWorksheet {
                wb: Arc::clone(&self.wb),
                sheet_title: Arc::new(RwLock::new(title)),
            })
            .collect()
    }

    #[pyo3(signature = (title=None, index=None))]
    pub fn create_sheet(&self, title: Option<&str>, index: Option<usize>) -> PyResult<PyWorksheet> {
        let mut wb = self.wb.write();
        let ws = wb.create_sheet(title, index).map_err(map_pyxl_err)?;
        let sheet_title = ws.title.clone();
        Ok(PyWorksheet {
            wb: Arc::clone(&self.wb),
            sheet_title: Arc::new(RwLock::new(sheet_title)),
        })
    }

    pub fn remove(&self, worksheet: &Bound<'_, PyWorksheet>) -> PyResult<()> {
        let ws = worksheet.borrow();
        let title = ws.sheet_title.read().clone();
        let mut wb = self.wb.write();
        wb.remove_sheet(&title).map_err(map_pyxl_err)?;
        Ok(())
    }

    pub fn remove_sheet(&self, title: &str) -> PyResult<()> {
        let mut wb = self.wb.write();
        wb.remove_sheet(title).map_err(map_pyxl_err)?;
        Ok(())
    }

    pub fn __getitem__(&self, title: &str) -> PyResult<PyWorksheet> {
        let wb = self.wb.read();
        if wb.contains_sheet(title) {
            Ok(PyWorksheet {
                wb: Arc::clone(&self.wb),
                sheet_title: Arc::new(RwLock::new(title.to_string())),
            })
        } else {
            Err(PyKeyError::new_err(format!(
                "Worksheet {} does not exist.",
                title
            )))
        }
    }

    pub fn __contains__(&self, title: &str) -> bool {
        self.wb.read().contains_sheet(title)
    }

    pub fn __delitem__(&self, title: &str) -> PyResult<()> {
        let mut wb = self.wb.write();
        wb.remove_sheet(title).map_err(map_pyxl_err)?;
        Ok(())
    }

    pub fn __len__(&self) -> usize {
        self.wb.read().sheet_count()
    }

    pub fn __iter__(slf: PyRef<'_, Self>) -> PyWorkbookIter {
        let wb = slf.wb.read();
        let titles = wb.sheetnames();
        PyWorkbookIter {
            wb: Arc::clone(&slf.wb),
            sheet_titles: titles,
            index: 0,
        }
    }

    pub fn close(&self) -> PyResult<()> {
        Ok(())
    }

    pub fn __enter__(slf: PyRef<'_, Self>) -> PyRef<'_, Self> {
        slf
    }

    #[pyo3(signature = (_exc_type=None, _exc_val=None, _exc_tb=None))]
    pub fn __exit__(
        &self,
        _exc_type: Option<&Bound<'_, PyAny>>,
        _exc_val: Option<&Bound<'_, PyAny>>,
        _exc_tb: Option<&Bound<'_, PyAny>>,
    ) -> PyResult<bool> {
        self.close()?;
        Ok(false)
    }

    pub fn save(&self, _path: &str) -> PyResult<()> {
        Err(PyNotImplementedError::new_err(
            "Workbook saving will be implemented in Phase 3 (High-Performance XLSX Writer)",
        ))
    }

    pub fn __repr__(&self) -> String {
        format!("<Workbook sheetnames={:?}>", self.wb.read().sheetnames())
    }
}

impl Default for PyWorkbook {
    fn default() -> Self {
        Self::new()
    }
}

#[pyclass]
pub struct PyWorkbookIter {
    wb: Arc<RwLock<Workbook>>,
    sheet_titles: Vec<String>,
    index: usize,
}

#[pymethods]
impl PyWorkbookIter {
    fn __iter__(slf: PyRef<'_, Self>) -> PyRef<'_, Self> {
        slf
    }

    fn __next__(&mut self) -> Option<PyWorksheet> {
        if self.index < self.sheet_titles.len() {
            let title = self.sheet_titles[self.index].clone();
            self.index += 1;
            Some(PyWorksheet {
                wb: Arc::clone(&self.wb),
                sheet_title: Arc::new(RwLock::new(title)),
            })
        } else {
            None
        }
    }
}

// -------------------------------------------------------------
// Read-Only Streaming Classes
// -------------------------------------------------------------

#[pyclass(name = "ReadOnlyCell")]
pub struct PyReadOnlyCell {
    #[pyo3(get)]
    pub row: u32,
    #[pyo3(get)]
    pub column: u32,
    coordinate: String,
    column_letter: String,
    value: PyObject,
    data_type: &'static str,
}

#[pymethods]
impl PyReadOnlyCell {
    #[getter]
    pub fn coordinate(&self) -> &str {
        &self.coordinate
    }

    #[getter]
    pub fn column_letter(&self) -> &str {
        &self.column_letter
    }

    #[getter]
    pub fn value<'py>(&self, py: Python<'py>) -> PyObject {
        self.value.clone_ref(py)
    }

    #[setter]
    pub fn set_value(&self, _val: &Bound<'_, PyAny>) -> PyResult<()> {
        Err(PyTypeError::new_err("Cannot assign to read-only cell"))
    }

    #[getter]
    pub fn data_type(&self) -> &'static str {
        self.data_type
    }

    pub fn __repr__(&self) -> String {
        format!("<ReadOnlyCell {}.{}>", self.coordinate, self.row)
    }
}

pub type StreamingRowMessage = PyxlResult<Option<Vec<(u32, CellValue)>>>;

#[pyclass(name = "RowIterator")]
pub struct PyRowIterator {
    rx: std::sync::mpsc::Receiver<StreamingRowMessage>,
    min_col: u32,
    max_col: u32,
    values_only: bool,
    data_only: bool,
    current_row_idx: u32,
}

#[pymethods]
impl PyRowIterator {
    fn __iter__(slf: PyRef<'_, Self>) -> PyRef<'_, Self> {
        slf
    }

    fn __next__<'py>(&mut self, py: Python<'py>) -> PyResult<Option<Bound<'py, PyTuple>>> {
        match self.rx.recv() {
            Ok(Ok(Some(cells))) => {
                let row_num = self.current_row_idx;
                self.current_row_idx += 1;

                let width = (self.max_col - self.min_col + 1) as usize;
                let mut cell_map = HashMap::with_capacity(cells.len());
                for (col, val) in cells {
                    cell_map.insert(col, val);
                }

                let mut row_items = Vec::with_capacity(width);
                if self.values_only {
                    for col in self.min_col..=self.max_col {
                        if let Some(val) = cell_map.remove(&col) {
                            let resolved = match val {
                                CellValue::Formula {
                                    cached_value,
                                    formula,
                                } => {
                                    if self.data_only {
                                        cached_value.map(|b| *b).unwrap_or(CellValue::Empty)
                                    } else {
                                        CellValue::Formula {
                                            cached_value,
                                            formula,
                                        }
                                    }
                                }
                                other => other,
                            };
                            row_items.push(cell_value_to_python(py, &resolved)?);
                        } else {
                            row_items.push(py.None().into_bound(py));
                        }
                    }
                } else {
                    for col in self.min_col..=self.max_col {
                        let coord = tuple_to_coordinate(row_num, col).map_err(map_pyxl_err)?;
                        let col_letter = get_column_letter(col).map_err(map_pyxl_err)?;
                        if let Some(val) = cell_map.remove(&col) {
                            let dt = val.data_type();
                            let resolved = match val {
                                CellValue::Formula {
                                    cached_value,
                                    formula,
                                } => {
                                    if self.data_only {
                                        cached_value.map(|b| *b).unwrap_or(CellValue::Empty)
                                    } else {
                                        CellValue::Formula {
                                            cached_value,
                                            formula,
                                        }
                                    }
                                }
                                other => other,
                            };
                            let py_val = cell_value_to_python(py, &resolved)?.unbind();
                            let cell = Bound::new(
                                py,
                                PyReadOnlyCell {
                                    row: row_num,
                                    column: col,
                                    coordinate: coord,
                                    column_letter: col_letter,
                                    value: py_val,
                                    data_type: dt,
                                },
                            )?;
                            row_items.push(cell.into_any());
                        } else {
                            let cell = Bound::new(
                                py,
                                PyReadOnlyCell {
                                    row: row_num,
                                    column: col,
                                    coordinate: coord,
                                    column_letter: col_letter,
                                    value: py.None(),
                                    data_type: "n",
                                },
                            )?;
                            row_items.push(cell.into_any());
                        }
                    }
                }
                Ok(Some(PyTuple::new_bound(py, row_items)))
            }
            Ok(Ok(None)) | Err(_) => Ok(None),
            Ok(Err(e)) => Err(map_pyxl_err(e)),
        }
    }
}

#[pyclass(name = "ReadOnlyWorksheet")]
pub struct PyReadOnlyWorksheet {
    file_path: String,
    title: String,
    target_path: String,
    shared_strings: Arc<SharedStringsTable>,
    styles: Arc<Styles>,
    min_row: u32,
    max_row: u32,
    min_col: u32,
    max_col: u32,
    data_only: bool,
}

#[pymethods]
impl PyReadOnlyWorksheet {
    #[getter]
    pub fn title(&self) -> &str {
        &self.title
    }

    #[getter]
    pub fn min_row(&self) -> u32 {
        self.min_row
    }

    #[getter]
    pub fn max_row(&self) -> u32 {
        self.max_row
    }

    #[getter]
    pub fn min_column(&self) -> u32 {
        self.min_col
    }

    #[getter]
    pub fn max_column(&self) -> u32 {
        self.max_col
    }

    #[getter]
    pub fn dimensions(&self) -> PyResult<String> {
        let start = tuple_to_coordinate(self.min_row, self.min_col).map_err(map_pyxl_err)?;
        let end = tuple_to_coordinate(self.max_row, self.max_col).map_err(map_pyxl_err)?;
        Ok(format!("{}:{}", start, end))
    }

    pub fn calculate_dimension(&self) -> PyResult<String> {
        self.dimensions()
    }

    pub fn cell<'py>(&self, py: Python<'py>, row: u32, column: u32) -> PyResult<PyReadOnlyCell> {
        let file = std::fs::File::open(&self.file_path)
            .map_err(|e| map_pyxl_err(PyxlError::IoError(format!("Failed to open file: {}", e))))?;
        let mut archive = zip::ZipArchive::new(std::io::BufReader::new(file))
            .map_err(|e| map_pyxl_err(PyxlError::IoError(format!("Invalid zip: {}", e))))?;
        let sheet_file = archive.by_name(&self.target_path).map_err(|e| {
            map_pyxl_err(PyxlError::IoError(format!("Cannot open sheet XML: {}", e)))
        })?;

        let cell_val = read_single_cell(
            sheet_file,
            row,
            column,
            &self.shared_strings,
            &self.styles,
            self.data_only,
        )
        .map_err(map_pyxl_err)?;

        let dt = cell_val.data_type();
        let py_val = cell_value_to_python(py, &cell_val)?.unbind();
        let coord = tuple_to_coordinate(row, column).map_err(map_pyxl_err)?;
        let col_letter = get_column_letter(column).map_err(map_pyxl_err)?;

        Ok(PyReadOnlyCell {
            row,
            column,
            coordinate: coord,
            column_letter: col_letter,
            value: py_val,
            data_type: dt,
        })
    }

    pub fn __getitem__<'py>(
        &self,
        py: Python<'py>,
        key: &Bound<'py, PyAny>,
    ) -> PyResult<Bound<'py, PyAny>> {
        let key_str = key.extract::<String>()?;
        if key_str.contains(':') {
            let (min_r, min_c, max_r, max_c) = range_boundaries(&key_str).map_err(map_pyxl_err)?;
            let mut outer = Vec::with_capacity((max_r - min_r + 1) as usize);
            for r in min_r..=max_r {
                let mut inner = Vec::with_capacity((max_c - min_c + 1) as usize);
                for c in min_c..=max_c {
                    let cell = self.cell(py, r, c)?;
                    inner.push(Bound::new(py, cell)?.into_any());
                }
                outer.push(PyTuple::new_bound(py, inner));
            }
            Ok(PyTuple::new_bound(py, outer).into_any())
        } else {
            let (r, c) = coordinate_to_tuple(&key_str).map_err(map_pyxl_err)?;
            let cell = self.cell(py, r, c)?;
            Ok(Bound::new(py, cell)?.into_any())
        }
    }

    pub fn __setitem__(&self, _key: &str, _value: &Bound<'_, PyAny>) -> PyResult<()> {
        Err(PyTypeError::new_err("Worksheet is read-only"))
    }

    pub fn append(&self, _iterable: &Bound<'_, PyAny>) -> PyResult<u32> {
        Err(PyTypeError::new_err("Worksheet is read-only"))
    }

    #[pyo3(signature = (min_row=None, max_row=None, min_col=None, max_col=None, values_only=false))]
    pub fn iter_rows(
        &self,
        min_row: Option<u32>,
        max_row: Option<u32>,
        min_col: Option<u32>,
        max_col: Option<u32>,
        values_only: bool,
    ) -> PyResult<PyRowIterator> {
        let r_start = min_row.unwrap_or(self.min_row);
        let r_end = max_row.unwrap_or(self.max_row);
        let c_start = min_col.unwrap_or(self.min_col);
        let c_end = max_col.unwrap_or(self.max_col);

        let (tx, rx) =
            std::sync::mpsc::sync_channel::<PyxlResult<Option<Vec<(u32, CellValue)>>>>(64);

        let file_path = self.file_path.clone();
        let target_path = self.target_path.clone();
        let shared_strings = Arc::clone(&self.shared_strings);
        let styles = Arc::clone(&self.styles);
        let data_only = self.data_only;

        std::thread::spawn(move || {
            let file = match std::fs::File::open(&file_path) {
                Ok(f) => f,
                Err(e) => {
                    let _ = tx.send(Err(PyxlError::IoError(format!(
                        "Failed to open {}: {}",
                        file_path, e
                    ))));
                    return;
                }
            };
            let mut archive: zip::ZipArchive<std::io::BufReader<std::fs::File>> =
                match zip::ZipArchive::new(std::io::BufReader::new(file)) {
                    Ok(a) => a,
                    Err(e) => {
                        let _ = tx.send(Err(PyxlError::IoError(format!("Invalid zip: {}", e))));
                        return;
                    }
                };
            let sheet_file = match archive.by_name(&target_path) {
                Ok(f) => f,
                Err(e) => {
                    let _ = tx.send(Err(PyxlError::IoError(format!(
                        "Cannot open {}: {}",
                        target_path, e
                    ))));
                    return;
                }
            };

            let mut reader = StreamingRowReader::new(
                sheet_file,
                shared_strings,
                styles,
                r_start,
                r_end,
                c_start,
                c_end,
                data_only,
            );

            let mut next_expected = r_start;
            loop {
                match reader.next_xml_row() {
                    Ok(Some((xml_row, cells))) => {
                        while next_expected < xml_row {
                            if tx.send(Ok(Some(Vec::new()))).is_err() {
                                return;
                            }
                            next_expected += 1;
                        }
                        if tx.send(Ok(Some(cells))).is_err() {
                            return;
                        }
                        next_expected = xml_row + 1;
                    }
                    Ok(None) => break,
                    Err(e) => {
                        let _ = tx.send(Err(e));
                        return;
                    }
                }
            }
            while next_expected <= r_end {
                if tx.send(Ok(Some(Vec::new()))).is_err() {
                    return;
                }
                next_expected += 1;
            }
            let _ = tx.send(Ok(None));
        });

        Ok(PyRowIterator {
            rx,
            min_col: c_start,
            max_col: c_end,
            values_only,
            data_only: self.data_only,
            current_row_idx: r_start,
        })
    }

    #[getter]
    pub fn values(&self) -> PyResult<PyRowIterator> {
        self.iter_rows(None, None, None, None, true)
    }

    #[getter]
    pub fn rows(&self) -> PyResult<PyRowIterator> {
        self.iter_rows(None, None, None, None, false)
    }

    pub fn __repr__(&self) -> String {
        format!("<ReadOnlyWorksheet \"{}\">", self.title)
    }
}

#[pyclass(name = "ReadOnlyWorkbook")]
pub struct PyReadOnlyWorkbook {
    file_path: String,
    sheet_infos: Vec<SheetInfo>,
    shared_strings: Arc<SharedStringsTable>,
    styles: Arc<Styles>,
    active_index: usize,
    data_only: bool,
    is_closed: bool,
}

#[pymethods]
impl PyReadOnlyWorkbook {
    #[getter]
    pub fn read_only(&self) -> bool {
        true
    }

    #[getter]
    pub fn data_only(&self) -> bool {
        self.data_only
    }

    #[getter]
    pub fn sheetnames(&self) -> Vec<String> {
        self.sheet_infos.iter().map(|s| s.name.clone()).collect()
    }

    #[getter]
    pub fn active(&self) -> PyResult<PyReadOnlyWorksheet> {
        if self.sheet_infos.is_empty() {
            return Err(PyKeyError::new_err("Workbook has no sheets"));
        }
        let idx = if self.active_index < self.sheet_infos.len() {
            self.active_index
        } else {
            0
        };
        let info = &self.sheet_infos[idx];
        Ok(PyReadOnlyWorksheet {
            file_path: self.file_path.clone(),
            title: info.name.clone(),
            target_path: info.target_path.clone(),
            shared_strings: Arc::clone(&self.shared_strings),
            styles: Arc::clone(&self.styles),
            min_row: info.min_row,
            max_row: info.max_row,
            min_col: info.min_col,
            max_col: info.max_col,
            data_only: self.data_only,
        })
    }

    #[getter]
    pub fn worksheets(&self) -> Vec<PyReadOnlyWorksheet> {
        self.sheet_infos
            .iter()
            .map(|info| PyReadOnlyWorksheet {
                file_path: self.file_path.clone(),
                title: info.name.clone(),
                target_path: info.target_path.clone(),
                shared_strings: Arc::clone(&self.shared_strings),
                styles: Arc::clone(&self.styles),
                min_row: info.min_row,
                max_row: info.max_row,
                min_col: info.min_col,
                max_col: info.max_col,
                data_only: self.data_only,
            })
            .collect()
    }

    pub fn __getitem__(&self, title: &str) -> PyResult<PyReadOnlyWorksheet> {
        let info = self
            .sheet_infos
            .iter()
            .find(|s| s.name == title)
            .ok_or_else(|| PyKeyError::new_err(format!("Worksheet {} does not exist.", title)))?;

        Ok(PyReadOnlyWorksheet {
            file_path: self.file_path.clone(),
            title: info.name.clone(),
            target_path: info.target_path.clone(),
            shared_strings: Arc::clone(&self.shared_strings),
            styles: Arc::clone(&self.styles),
            min_row: info.min_row,
            max_row: info.max_row,
            min_col: info.min_col,
            max_col: info.max_col,
            data_only: self.data_only,
        })
    }

    pub fn __contains__(&self, title: &str) -> bool {
        self.sheet_infos.iter().any(|s| s.name == title)
    }

    pub fn __len__(&self) -> usize {
        self.sheet_infos.len()
    }

    pub fn __iter__(&self) -> PyReadOnlyWorkbookIter {
        PyReadOnlyWorkbookIter {
            file_path: self.file_path.clone(),
            sheet_infos: self.sheet_infos.clone(),
            shared_strings: Arc::clone(&self.shared_strings),
            styles: Arc::clone(&self.styles),
            data_only: self.data_only,
            index: 0,
        }
    }

    pub fn close(&mut self) -> PyResult<()> {
        self.is_closed = true;
        Ok(())
    }

    pub fn __enter__(slf: PyRef<'_, Self>) -> PyRef<'_, Self> {
        slf
    }

    #[pyo3(signature = (_exc_type=None, _exc_val=None, _exc_tb=None))]
    pub fn __exit__(
        &mut self,
        _exc_type: Option<&Bound<'_, PyAny>>,
        _exc_val: Option<&Bound<'_, PyAny>>,
        _exc_tb: Option<&Bound<'_, PyAny>>,
    ) -> PyResult<bool> {
        self.close()?;
        Ok(false)
    }

    pub fn __repr__(&self) -> String {
        format!("<ReadOnlyWorkbook sheetnames={:?}>", self.sheetnames())
    }
}

#[pyclass]
pub struct PyReadOnlyWorkbookIter {
    file_path: String,
    sheet_infos: Vec<SheetInfo>,
    shared_strings: Arc<SharedStringsTable>,
    styles: Arc<Styles>,
    data_only: bool,
    index: usize,
}

#[pymethods]
impl PyReadOnlyWorkbookIter {
    fn __iter__(slf: PyRef<'_, Self>) -> PyRef<'_, Self> {
        slf
    }

    fn __next__(&mut self) -> Option<PyReadOnlyWorksheet> {
        if self.index < self.sheet_infos.len() {
            let info = &self.sheet_infos[self.index];
            self.index += 1;
            Some(PyReadOnlyWorksheet {
                file_path: self.file_path.clone(),
                title: info.name.clone(),
                target_path: info.target_path.clone(),
                shared_strings: Arc::clone(&self.shared_strings),
                styles: Arc::clone(&self.styles),
                min_row: info.min_row,
                max_row: info.max_row,
                min_col: info.min_col,
                max_col: info.max_col,
                data_only: self.data_only,
            })
        } else {
            None
        }
    }
}

// -------------------------------------------------------------
// Public load_workbook binding
// -------------------------------------------------------------

#[pyfunction]
#[pyo3(signature = (path, read_only=false, data_only=false))]
pub fn load_workbook<'py>(
    py: Python<'py>,
    path: &str,
    read_only: bool,
    data_only: bool,
) -> PyResult<Bound<'py, PyAny>> {
    let p = std::path::Path::new(path);
    if !p.exists() {
        return Err(pyo3::exceptions::PyFileNotFoundError::new_err(format!(
            "No such file or directory: '{}'",
            path
        )));
    }

    let mut package = XlsxPackage::open(path, read_only, data_only).map_err(map_pyxl_err)?;

    if read_only {
        let ro_wb = PyReadOnlyWorkbook {
            file_path: path.to_string(),
            sheet_infos: package.sheets,
            shared_strings: package.shared_strings,
            styles: package.styles,
            active_index: package.active_sheet_index,
            data_only,
            is_closed: false,
        };
        Ok(Bound::new(py, ro_wb)?.into_any())
    } else {
        let wb = package.to_workbook().map_err(map_pyxl_err)?;
        let py_wb = PyWorkbook {
            wb: Arc::new(RwLock::new(wb)),
        };
        Ok(Bound::new(py, py_wb)?.into_any())
    }
}
