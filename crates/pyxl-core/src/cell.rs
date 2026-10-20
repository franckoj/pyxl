use crate::coordinate::{get_column_letter, tuple_to_coordinate};
use crate::error::PyxlResult;

#[derive(Debug, Clone, PartialEq, Default)]
pub enum CellValue {
    #[default]
    Empty,
    String(String),
    Int(i64),
    Float(f64),
    Bool(bool),
    DateTime(String),
    Formula(String),
}

impl CellValue {
    pub fn data_type(&self) -> &'static str {
        match self {
            CellValue::Empty => "n",
            CellValue::String(_) => "s",
            CellValue::Int(_) | CellValue::Float(_) => "n",
            CellValue::Bool(_) => "b",
            CellValue::DateTime(_) => "d",
            CellValue::Formula(_) => "f",
        }
    }

    pub fn is_empty(&self) -> bool {
        matches!(self, CellValue::Empty)
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Cell {
    pub row: u32,
    pub column: u32,
    pub value: CellValue,
}

impl Cell {
    pub fn new(row: u32, column: u32, value: CellValue) -> Self {
        Self { row, column, value }
    }

    pub fn coordinate(&self) -> PyxlResult<String> {
        tuple_to_coordinate(self.row, self.column)
    }

    pub fn column_letter(&self) -> PyxlResult<String> {
        get_column_letter(self.column)
    }

    pub fn data_type(&self) -> &'static str {
        self.value.data_type()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_cell_creation_and_properties() {
        let cell = Cell::new(1, 1, CellValue::String("hello".to_string()));
        assert_eq!(cell.row, 1);
        assert_eq!(cell.column, 1);
        assert_eq!(cell.coordinate().unwrap(), "A1");
        assert_eq!(cell.column_letter().unwrap(), "A");
        assert_eq!(cell.data_type(), "s");

        let num_cell = Cell::new(5, 27, CellValue::Int(42));
        assert_eq!(num_cell.coordinate().unwrap(), "AA5");
        assert_eq!(num_cell.data_type(), "n");
    }
}
