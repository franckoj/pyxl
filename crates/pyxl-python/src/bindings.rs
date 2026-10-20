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
use pyxl_core::error::PyxlError;
use pyxl_core::workbook::Workbook;

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
            Ok(CellValue::Formula(s))
        } else {
            Ok(CellValue::String(s))
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
        CellValue::DateTime(d) => Ok(d.to_object(py).into_bound(py)),
        CellValue::Formula(f) => Ok(f.to_object(py).into_bound(py)),
    }
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
        let title = self.sheet_title.read().clone();
        let ws = wb
            .get_sheet(&title)
            .ok_or_else(|| PyxlError::SheetNotFound(title.clone()))
            .map_err(map_pyxl_err)?;
        let val = ws.get_cell_value(self.row, self.column);
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
            // Range access e.g. "A1:B10"
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
            // Single cell access e.g. "A1"
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
            let title = self.sheet_title.read().clone();
            let ws = wb
                .get_sheet(&title)
                .ok_or_else(|| PyxlError::SheetNotFound(title.clone()))
                .map_err(map_pyxl_err)?;

            for r in r_start..=r_end {
                let mut row_vals = Vec::with_capacity((c_end - c_start + 1) as usize);
                for c in c_start..=c_end {
                    let val = ws.get_cell_value(r, c);
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
            let title = self.sheet_title.read().clone();
            let ws = wb
                .get_sheet(&title)
                .ok_or_else(|| PyxlError::SheetNotFound(title.clone()))
                .map_err(map_pyxl_err)?;

            for c in c_start..=c_end {
                let mut col_vals = Vec::with_capacity((r_end - r_start + 1) as usize);
                for r in r_start..=r_end {
                    let val = ws.get_cell_value(r, c);
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
        // In-memory workbook close is a clean no-op
        Ok(())
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
