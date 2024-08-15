use regex::Regex;

#[derive(Debug, Clone, PartialEq)]
pub enum BusinessType {
    Unclassified,    // 미정
    GeneralBusiness, // 일반사업
    PledgedBusiness, // 공약사업
    OngoingBusiness, // 상시사업
}

impl Default for BusinessType {
    fn default() -> Self {
        BusinessType::Unclassified
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Date {
    pub year: u16,
    pub month: u8,
    pub day: u8,
}

impl Date {
    pub fn new(date: &str) -> Date {
        let caps = Regex::new(r"(\d{4})\.(\d{2})\.(\d{2})")
            .unwrap()
            .captures(date)
            .unwrap();

        Date {
            year: caps[1].parse().unwrap(),
            month: caps[2].parse().unwrap(),
            day: caps[3].parse().unwrap(),
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Data {
    pub date: Date,
    pub business_type: BusinessType,
    pub business_name: Option<String>,
    pub cash_in: u32,
    pub cash_out: u32,
    pub balance: u32,
    pub remarks: Option<String>,
    pub receipt_num: Option<String>,
}

#[derive(Default)]
pub struct DataBuilder {
    date: Option<Date>,
    business_type: BusinessType,
    business_name: Option<String>,
    cash_in: Option<u32>,
    cash_out: Option<u32>,
    balance: Option<u32>,
    remarks: Option<String>,
    receipt_num: Option<String>,
}

impl DataBuilder {
    pub fn new() -> Self {
        DataBuilder::default()
    }
    pub fn date(&mut self, date: Date) -> &mut Self {
        self.date = Some(date);
        self
    }
    pub fn business_type(&mut self, business_type: BusinessType) -> &mut Self {
        self.business_type = business_type;
        self
    }
    pub fn business_name(&mut self, business_name: impl Into<String>) -> &mut Self {
        self.business_name = Some(business_name.into());
        self
    }
    pub fn cash_in(&mut self, cash_in: u32) -> &mut Self {
        self.cash_in = Some(cash_in);
        self
    }
    pub fn cash_out(&mut self, cash_out: u32) -> &mut Self {
        self.cash_out = Some(cash_out);
        self
    }
    pub fn balance(&mut self, balance: u32) -> &mut Self {
        self.balance = Some(balance);
        self
    }
    pub fn remarks(&mut self, remarks: impl Into<String>) -> &mut Self {
        self.remarks = Some(remarks.into());
        self
    }
    pub fn receipt_num(&mut self, receipt_num: impl Into<String>) -> &mut Self {
        self.receipt_num = Some(receipt_num.into());
        self
    }
    pub fn build(&self) -> Result<Data, &'static str> {
        let date = self.date.clone().ok_or("No DATE provided")?;
        let cash_in = self.cash_in.ok_or("No CASH_IN provided")?;
        let cash_out = self.cash_out.ok_or("No CASH_OUT provided")?;
        let balance = self.balance.ok_or("No BALANCE provided")?;
        Ok(Data {
            date,
            business_type: self.business_type.clone(),
            business_name: self.business_name.clone(),
            cash_in,
            cash_out,
            balance,
            remarks: self.remarks.clone(),
            receipt_num: self.receipt_num.clone(),
        })
    }
}
