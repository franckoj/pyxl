//! Shared strings table handling for OOXML.

#[derive(Debug, Default, Clone)]
pub struct SharedStringsTable {
    strings: Vec<String>,
}

impl SharedStringsTable {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn get(&self, idx: usize) -> Option<&str> {
        self.strings.get(idx).map(|s| s.as_str())
    }

    pub fn len(&self) -> usize {
        self.strings.len()
    }

    pub fn is_empty(&self) -> bool {
        self.strings.is_empty()
    }
}
