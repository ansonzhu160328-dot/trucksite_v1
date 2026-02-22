from pydantic import BaseModel, Field

class CalcRequest(BaseModel):
    site_location: str = Field("", description="场站位置/地区")

    site_length_m: float = Field(..., ge=0)
    site_width_m: float = Field(..., ge=0)
    

    pile_kva_per: float = Field(400.0, gt=0)
    guns_per_pile: int = Field(2, ge=1, le=8)

    kwh_per_gun_per_day: float = Field(1000.0, gt=0)
    service_fee_yuan_per_kwh: float = Field(0.3, ge=0)
    days_per_year: int = Field(330, ge=1, le=366)

    power_cost_yuan_per_kva: float = Field(600.0, ge=0)
    civil_cost_yuan_per_sqm: float = Field(200.0, ge=0)
    pile_cost_yuan_each: float = Field(45000.0, ge=0)


    rent_yuan_per_sqm_month: float = Field(0.0, ge=0)
    staff_count: int = Field(0, ge=0)
    salary_yuan_per_month: float = Field(0.0, ge=0)
