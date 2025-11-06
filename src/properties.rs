// Properties module for document metadata
use napi_derive::napi;
use rust_xlsxwriter::DocProperties;

#[napi]
pub struct RuscDocProperties {
    pub(crate) properties: DocProperties,
}

#[napi]
impl RuscDocProperties {
    /// Create a new DocProperties struct
    #[napi(factory)]
    pub fn new() -> Self {
        RuscDocProperties {
            properties: DocProperties::new(),
        }
    }

    /// Set the document author
    #[napi]
    pub fn set_author(&mut self, author: String) -> &Self {
        self.properties = self.properties.clone().set_author(&author);
        self
    }

    /// Set the document title
    #[napi]
    pub fn set_title(&mut self, title: String) -> &Self {
        self.properties = self.properties.clone().set_title(&title);
        self
    }

    /// Set the document subject
    #[napi]
    pub fn set_subject(&mut self, subject: String) -> &Self {
        self.properties = self.properties.clone().set_subject(&subject);
        self
    }

    /// Set the document comment
    #[napi]
    pub fn set_comment(&mut self, comment: String) -> &Self {
        self.properties = self.properties.clone().set_comment(&comment);
        self
    }

    /// Set the document company
    #[napi]
    pub fn set_company(&mut self, company: String) -> &Self {
        self.properties = self.properties.clone().set_company(&company);
        self
    }

    /// Set the document manager
    #[napi]
    pub fn set_manager(&mut self, manager: String) -> &Self {
        self.properties = self.properties.clone().set_manager(&manager);
        self
    }

    /// Set the document status
    #[napi]
    pub fn set_status(&mut self, status: String) -> &Self {
        self.properties = self.properties.clone().set_status(&status);
        self
    }

    /// Set the document category
    #[napi]
    pub fn set_category(&mut self, category: String) -> &Self {
        self.properties = self.properties.clone().set_category(&category);
        self
    }

    /// Set the document keywords
    #[napi]
    pub fn set_keywords(&mut self, keywords: String) -> &Self {
        self.properties = self.properties.clone().set_keywords(&keywords);
        self
    }

    /// Set the document hyperlink base
    #[napi]
    pub fn set_hyperlink_base(&mut self, base: String) -> &Self {
        self.properties = self.properties.clone().set_hyperlink_base(&base);
        self
    }

    /// Get the internal properties reference
    pub(crate) fn get_properties(&self) -> &DocProperties {
        &self.properties
    }
}
