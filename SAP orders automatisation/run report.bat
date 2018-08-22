::[Bat To Exe Converter]
::
::YAwzoRdxOk+EWAjk
::fBw5plQjdCyDJGyX8VAjFDJgfy24HUaGIrAP4/z0/9aOoUITGus8d+8=
::YAwzuBVtJxjWCl3EqQJgSA==
::ZR4luwNxJguZRRnk
::Yhs/ulQjdF+5
::cxAkpRVqdFKZSDk=
::cBs/ulQjdF+5
::ZR41oxFsdFKZSDk=
::eBoioBt6dFKZSDk=
::cRo6pxp7LAbNWATEpCI=
::egkzugNsPRvcWATEpCI=
::dAsiuh18IRvcCxnZtBJQ
::cRYluBh/LU+EWAnk
::YxY4rhs+aU+IeA==
::cxY6rQJ7JhzQF1fEqQJhZksaHErSXA==
::ZQ05rAF9IBncCkqN+0xwdVsFAlTMbCXqZg==
::ZQ05rAF9IAHYFVzEqQIRHDd1f0SmPWK2H/UT+uz+/fnHpUgTUfA+bIDJug==
::eg0/rx1wNQPfEVWB+kM9LVsJDGQ=
::fBEirQZwNQPfEVWB+kM9LVsJDGQ=
::cRolqwZ3JBvQF1fEqQIcKQ5aTwyHLyu+B7wQ8ajp6vqIsFldU+cxfZ3azrucQA==
::dhA7uBVwLU+EWHim1iI=
::YQ03rBFzNR3SWATElA==
::dhAmsQZ3MwfNWATElA==
::ZQ0/vhVqMQ3MEVWAtB9wSA==
::Zg8zqx1/OA3MEVWAtB9wSA==
::dhA7pRFwIByZRRnk
::Zh4grVQjdD6DJGq2wn0RBSd1b0mhGFeOIpcgzs3I07jX8xghcdJ/VY7J0bGaKe4UqkTqcdYe13Zfi/crABRafx6XTzsYiF0CkmWMO97cnB3lT1qa2hlgSj1IhHHVjT8+IMFtiswRx2675Eif
::YB416Ek+ZG8=
::
::
::978f952a14a936cc963da21a135fa983
@echo off
TITLE Daily orders generator
SETLOCAL ENABLEEXTENSIONS

if exist template.xlsb (
    echo Starting the report generator...
	start template.xlsb /popup
    echo.
    echo This window will automatically close in 10 seconds.. 
) else (
echo Error: Template file does not exist, please check the instructions...
)
timeout 10 >nul
exit