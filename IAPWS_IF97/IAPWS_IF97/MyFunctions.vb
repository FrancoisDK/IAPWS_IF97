
Imports ExcelDna.Integration
Imports ExcelDna.IntelliSense

Public Module MyFunctions

    <ExcelFunction(Description:="Density of water or steam with pressure and temperature known, in kg/m³", Category:="IAPWS-IF97")>
    Public Function WaterDensity(<ExcelArgument(Description:="Temperature in Kelvin")> Temperature As Double, <ExcelArgument(Description:="Pressure in bara")> Pressure As Double) As Double
        Dim Density As Double

        Density = densW(Temperature, Pressure)
        Return Density
    End Function


    <ExcelFunction(Description:="Density of saturated water with pressure known, in kg/m³", Category:="IAPWS-IF97")>
    Public Function DensSatLiqPWater(<ExcelArgument(Description:="Pressure in bara")> Pressure As Double) As Double
        Dim SatDensity As Double

        SatDensity = densSatLiqPW(Pressure)
        Return SatDensity

    End Function

    <ExcelFunction(Description:="Density of saturated water with temperature known, in kg/m³", Category:="IAPWS-IF97")>
    Public Function DensSatLiqTWater(<ExcelArgument(Description:="Temperature in Kelvin")> Temperature As Double) As Double
        Dim SatDensityLiq As Double

        SatDensityLiq = densSatLiqTW(Temperature)
        Return SatDensityLiq

    End Function


    <ExcelFunction(Description:="Density of saturated steam with pressure known, in kg/m³", Category:="IAPWS-IF97")>
    Public Function DensSatVapPWater(<ExcelArgument(Description:="Pressure in bara")> Pressure As Double) As Double
        Dim SatDensityVap As Double

        SatDensityVap = densSatVapPW(Pressure)
        Return SatDensityVap

    End Function
    <ExcelFunction(Description:="Density of saturated steam with temperature known, in kg/m³", Category:="IAPWS-IF97")>
    Public Function DensSatVapTWater(<ExcelArgument(Description:="Temperature in Kelvin")> Temperature As Double) As Double
        Dim SatDensityVap As Double

        SatDensityVap = densSatVapTW(Temperature)
        Return SatDensityVap

    End Function


    <ExcelFunction(Description:="Saturation pressure of water with temperature known, in bara", Category:="IAPWS-IF97")>
    Public Function PsatWater(<ExcelArgument(Description:="Temperature in Kelvin")> Temperature As Double) As Double
        Dim SatPressure As Double

        SatPressure = pSatW(Temperature)
        Return SatPressure

    End Function

    <ExcelFunction(Description:="Saturation temperature of water", Category:="IAPWS-IF97")>
    Public Function TsatWater(<ExcelArgument(Description:="Pressure in bara")> Pressure As Double) As Double
        Dim SatTemperature As Double

        SatTemperature = tSatW(Pressure)
        Return SatTemperature

    End Function

    <ExcelFunction(Description:="Specific internal energy of water, in kJ/kg", Category:="IAPWS-IF97")>
    Public Function EnergyWater(<ExcelArgument(Description:="Temperature in Kelvin")> Temperature As Double, <ExcelArgument(Description:="Pressure in bara")> Pressure As Double) As Double
        '
        ' specific internal energy of water or steam
        ' energyW in kJ/kg
        ' temperature in K
        ' pressure in bar
        '
        ' energyW = -1: temperature and/or pressure outside range

        Dim EWater As Double

        EWater = energyW(Temperature, Pressure)
        Return EWater

    End Function

    <ExcelFunction(Description:="Specific internal energy of saturated water at specified pressure, in kJ/kg", Category:="IAPWS-IF97")>
    Public Function EnergySatLiqPwater(<ExcelArgument(Description:="Pressure in bara")> Pressure As Double) As Double
        '
        ' specific internal energy of water or steam
        ' energyW in kJ/kg
        ' temperature in K
        ' pressure in bar
        '
        ' energyW = -1: temperature and/or pressure outside range

        Dim EWater As Double

        EWater = energySatLiqPW(Pressure)
        Return EWater

    End Function


    <ExcelFunction(Description:="Specific internal energy of saturated water at specified temperature, in kJ/kg", Category:="IAPWS-IF97")>
    Public Function EnergySatLiqTwater(<ExcelArgument(Description:="Temperature in Kelvin")> Temperature As Double) As Double
        '
        ' specific internal energy of water or steam
        ' energyW in kJ/kg
        ' temperature in K
        ' pressure in bar
        '
        ' energyW = -1: temperature and/or pressure outside range

        Dim EWater As Double

        EWater = energySatLiqTW(Temperature)
        Return EWater

    End Function

    '''''
    '''
    <ExcelFunction(Description:="Specific internal energy of saturated steam at specified pressure, in kJ/kg", Category:="IAPWS-IF97")>
    Public Function EnergySatVapPwater(<ExcelArgument(Description:="Pressure in bara")> Pressure As Double) As Double

        Dim EWater As Double

        EWater = energySatVapPW(Pressure)
        Return EWater

    End Function


    <ExcelFunction(Description:="Specific internal energy of saturated steam at specified temperature, in kJ/kg", Category:="IAPWS-IF97")>
    Public Function EnergySatVapTwater(<ExcelArgument(Description:="Temperature in Kelvin")> Temperature As Double) As Double


        Dim EWater As Double

        EWater = energySatVapTW(Temperature)
        Return EWater

    End Function

    '#########################################################
    <ExcelFunction(Description:="Specific entropy of water or steam at specified temperature and pressure, in kJ/kg,K", Category:="IAPWS-IF97")>
    Public Function EntropyWater(<ExcelArgument(Description:="Temperature in Kelvin")> Temperature As Double, <ExcelArgument(Description:="Pressure in bara")> Pressure As Double) As Double

        Dim EWater As Double

        EWater = entropyW(Temperature, Pressure)
        Return EWater

    End Function


    <ExcelFunction(Description:="Specific entropy of saturated water at specified pressure, in kJ/kg,K", Category:="IAPWS-IF97")>
    Public Function EntropySatLiqPwater(<ExcelArgument(Description:="Pressure in bara")> Pressure As Double) As Double
        '
        '
        ' specific entropy of saturated liquid water as a function of temperature
        ' entropySatLiqTW in kJ/(kg K)
        ' temperature in K
        '
        ' entropySatLiqTW = -1: temperature outside range
        Dim EWater As Double

        EWater = entropySatLiqPW(Pressure)
        Return EWater

    End Function


    <ExcelFunction(Description:="Specific entropy of saturated water at specified temperature, in kJ/kg,K", Category:="IAPWS-IF97")>
    Public Function EntropySatLiqTwater(<ExcelArgument(Description:="Temperature in Kelvin")> Temperature As Double) As Double

        Dim EWater As Double

        EWater = entropySatLiqTW(Temperature)
        Return EWater

    End Function

    '''''
    '''
    <ExcelFunction(Description:="Specific entropy of saturated steam at specified pressure, in kJ/kg,K", Category:="IAPWS-IF97")>
    Public Function EntropySatVapPwater(<ExcelArgument(Description:="Pressure in bara")> Pressure As Double) As Double

        Dim EWater As Double

        EWater = entropySatVapPW(Pressure)
        Return EWater

    End Function

    <ExcelFunction(Description:="Specific entropy of saturated steam at specified temperature, in kJ/kg,K", Category:="IAPWS-IF97")>
    Public Function EntropySatVapTwater(<ExcelArgument(Description:="Temperature in Kelvin")> Temperature As Double) As Double


        Dim EWater As Double

        EWater = entropySatVapTW(Temperature)
        Return EWater

    End Function
    '#######################################

    <ExcelFunction(Description:="Specific enthalpy of water or steam at specified temperature and pressure, in kJ/kg", Category:="IAPWS-IF97")>
    Public Function EnthalpyW(<ExcelArgument(Description:="Temperature in Kelvin")> Temperature As Double, <ExcelArgument(Description:="Pressure in bara")> Pressure As Double) As Double
        ' specific enthalpy of saturated liquid water as a function of temperature
        ' enthalpySatLiqTW in kJ/kg
        ' temperature in K
        '
        ' enthalpySatLiqTW = -1: temperature outside range
        Dim EWater As Double

        EWater = EnthalpyW(Temperature, Pressure)
        Return EWater

    End Function


    <ExcelFunction(Description:="Specific enthalpy of saturated water at specified pressure, in kJ/kg", Category:="IAPWS-IF97")>
    Public Function EnthalpySatLiqPwater(<ExcelArgument(Description:="Pressure in bara")> Pressure As Double) As Double
        '
        ' specific enthalpy of saturated liquid water as a function of temperature
        ' enthalpySatLiqTW in kJ/kg
        ' temperature in K
        '
        ' enthalpySatLiqTW = -1: temperature outside range
        Dim EWater As Double

        EWater = enthalpySatLiqPW(Pressure)
        Return EWater

    End Function


    <ExcelFunction(Description:="Specific enthalpy of saturated water at specified temperature, in kJ/kg", Category:="IAPWS-IF97")>
    Public Function EnthalpySatLiqTwater(<ExcelArgument(Description:="Temperature in Kelvin")> Temperature As Double) As Double

        Dim EWater As Double

        EWater = enthalpySatLiqTW(Temperature)
        Return EWater

    End Function

    '''''
    '''
    <ExcelFunction(Description:="Specific enthalpy of saturated steam at specified pressure, in kJ/kg", Category:="IAPWS-IF97")>
    Public Function EnthalpySatVapPwater(<ExcelArgument(Description:="Pressure in bara")> Pressure As Double) As Double

        Dim EWater As Double

        EWater = enthalpySatVapPW(Pressure)
        Return EWater

    End Function


    <ExcelFunction(Description:="Specific enthlapy of saturated steam at specified temperature, in kJ/kg", Category:="IAPWS-IF97")>
    Public Function EnthalpySatVapTwater(<ExcelArgument(Description:="Temperature in Kelvin")> Temperature As Double) As Double


        Dim EWater As Double

        EWater = enthalpySatVapTW(Temperature)
        Return EWater

    End Function

    '#################################

    <ExcelFunction(Description:="Specific isobaric heat capacity of saturated water at specified pressure, in kJ/kg,K", Category:="IAPWS-IF97")>
    Public Function CpWater(<ExcelArgument(Description:="Temperature in Kelvin")> Temperature As Double, <ExcelArgument(Description:="Pressure in bara")> Pressure As Double) As Double
        '
        ' specific isobaric heat capacity of saturated liquid water as a function of temperature
        ' cpSatLiqTW in kJ/(kg K)
        ' temperature in K
        '
        ' cpSatLiqTW = -1: temperature outside range
        Dim Cp As Double

        Cp = cpW(Temperature, Pressure)
        Return Cp

    End Function
    <ExcelFunction(Description:="Specific isobaric heat capacity of saturated water at specified pressure, in kJ/kg,K", Category:="IAPWS-IF97")>
    Public Function CpSatLiqPwater(<ExcelArgument(Description:="Pressure in bara")> Pressure As Double) As Double
        '
        ' specific isobaric heat capacity of saturated liquid water as a function of temperature
        ' cpSatLiqTW in kJ/(kg K)
        ' temperature in K
        '
        ' cpSatLiqTW = -1: temperature outside range
        Dim Cp As Double

        Cp = cpSatLiqPW(Pressure)
        Return Cp

    End Function


    <ExcelFunction(Description:="Specific isobaric heat capacity of saturated water at specified temperature, in kJ/kg,K", Category:="IAPWS-IF97")>
    Public Function CpSatLiqTwater(<ExcelArgument(Description:="Temperature in Kelvin")> Temperature As Double) As Double

        Dim Cp As Double

        Cp = cpSatLiqTW(Temperature)
        Return Cp

    End Function

    '''''
    '''
    <ExcelFunction(Description:="Specific isobaric heat capacity of saturated steam at specified pressure, in kJ/kg,K", Category:="IAPWS-IF97")>
    Public Function CpSatVapPwater(<ExcelArgument(Description:="Pressure in bara")> Pressure As Double) As Double

        Dim Cp As Double

        Cp = cpSatVapPW(Pressure)
        Return Cp

    End Function


    <ExcelFunction(Description:="Specific isobaric heat capacity of saturated steam at specified temperature, in kJ/kg,K", Category:="IAPWS-IF97")>
    Public Function CpSatVapTwater(<ExcelArgument(Description:="Temperature in Kelvin")> Temperature As Double) As Double


        Dim Cp As Double

        Cp = cpSatVapTW(Temperature)
        Return Cp

    End Function

    '#################################

    <ExcelFunction(Description:="Specific isochoric heat capacity of saturated water at specified pressure, in kJ/kg,K", Category:="IAPWS-IF97")>
    Public Function CvWater(<ExcelArgument(Description:="Temperature in Kelvin")> Temperature As Double, <ExcelArgument(Description:="Pressure in bara")> Pressure As Double) As Double
        '
        ' specific isochoric heat capacity of saturated liquid water as a function of temperature
        ' cpSatLiqTW in kJ/(kg K)
        ' temperature in K
        '
        ' cpSatLiqTW = -1: temperature outside range
        Dim Cv As Double

        Cv = cvW(Temperature, Pressure)
        Return Cv

    End Function
    <ExcelFunction(Description:="Specific isochoric  heat capacity of saturated water at specified pressure, in kJ/kg,K", Category:="IAPWS-IF97")>
    Public Function CvSatLiqPwater(<ExcelArgument(Description:="Pressure in bara")> Pressure As Double) As Double
        '
        ' specific isobaric heat capacity of saturated liquid water as a function of temperature
        ' cpSatLiqTW in kJ/(kg K)
        ' temperature in K
        '
        ' cpSatLiqTW = -1: temperature outside range
        Dim Cv As Double

        Cv = cvSatLiqPW(Pressure)
        Return Cv

    End Function


    <ExcelFunction(Description:="Specific isochoric  heat capacity of saturated water at specified temperature, in kJ/kg,K", Category:="IAPWS-IF97")>
    Public Function CvSatLiqTwater(<ExcelArgument(Description:="Temperature in Kelvin")> Temperature As Double) As Double

        Dim Cv As Double

        Cv = cvSatLiqTW(Temperature)
        Return Cv

    End Function

    '''''
    '''
    <ExcelFunction(Description:="Specific isochoric  heat capacity of saturated steam at specified pressure, in kJ/kg,K", Category:="IAPWS-IF97")>
    Public Function CvSatVapPwater(<ExcelArgument(Description:="Pressure in bara")> Pressure As Double) As Double

        Dim Cv As Double

        Cv = cvSatVapPW(Pressure)
        Return Cv

    End Function


    <ExcelFunction(Description:="Specific isochoric  heat capacity of saturated steam at specified temperature, in kJ/kg,K", Category:="IAPWS-IF97")>
    Public Function CvSatVapTwater(<ExcelArgument(Description:="Temperature in Kelvin")> Temperature As Double) As Double


        Dim Cv As Double

        Cv = cvSatVapTW(Temperature)
        Return Cv

    End Function

    '#################################

    <ExcelFunction(Description:="Viscocity of water or steam, in cP", Category:="IAPWS-IF97")>
    Public Function ViscWater(<ExcelArgument(Description:="Temperature in Kelvin")> Temperature As Double, <ExcelArgument(Description:="Pressure in bara")> Pressure As Double) As Double

        Dim Visc As Double

        Visc = viscW(Temperature, Pressure) * 1000
        Return Visc

    End Function
    <ExcelFunction(Description:="Viscocity of saturated water at specified pressure, in cP", Category:="IAPWS-IF97")>
    Public Function ViscSatLiqPwater(<ExcelArgument(Description:="Pressure in bara")> Pressure As Double) As Double

        Dim Visc As Double

        Visc = viscSatLiqPW(Pressure) * 1000
        Return Visc

    End Function


    <ExcelFunction(Description:="Viscocity of saturated water at specified temperature, in cP", Category:="IAPWS-IF97")>
    Public Function ViscSatLiqTwater(<ExcelArgument(Description:="Temperature in Kelvin")> Temperature As Double) As Double

        Dim Visc As Double

        Visc = viscSatLiqTW(Temperature) * 1000
        Return Visc

    End Function

    <ExcelFunction(Description:="Viscocity of saturated steam at specified pressure, in cP", Category:="IAPWS-IF97")>
    Public Function ViscSatVapPwater(<ExcelArgument(Description:="Pressure in bara")> Pressure As Double) As Double

        Dim Visc As Double

        Visc = viscSatVapPW(Pressure) * 1000
        Return Visc

    End Function


    <ExcelFunction(Description:="Viscocity of saturated steam at specified temperature, in cP", Category:="IAPWS-IF97")>
    Public Function ViscSatVapTwater(<ExcelArgument(Description:="Temperature in Kelvin")> Temperature As Double) As Double


        Dim Visc As Double

        Visc = viscSatVapTW(Temperature) * 1000
        Return Visc

    End Function

    '#################################

    <ExcelFunction(Description:="Thermal conductivity of water or steam, in W/m,K", Category:="IAPWS-IF97")>
    Public Function ThconWater(<ExcelArgument(Description:="Temperature in Kelvin")> Temperature As Double, <ExcelArgument(Description:="Pressure in bara")> Pressure As Double) As Double
        '
        ' Thermal conductivity of saturated liquid water as a function of temperature
        ' thconSatLiqTW in W /(m K)
        ' temperature in K
        '
        ' thconSatLiqTW = -1: temperature outside range
        '
        Dim Thcon As Double

        Thcon = thconW(Temperature, Pressure)
        Return Thcon

    End Function
    <ExcelFunction(Description:="Thermal conductivity of  saturated water at specified pressure, in W/m,K", Category:="IAPWS-IF97")>
    Public Function ThconSatLiqPwater(<ExcelArgument(Description:="Pressure in bara")> Pressure As Double) As Double

        Dim Thcon As Double

        Thcon = thconSatLiqPW(Pressure)
        Return Thcon

    End Function


    <ExcelFunction(Description:="Thermal conductivity of saturated water at specified temperature, in W/m,K", Category:="IAPWS-IF97")>
    Public Function ThconSatLiqTwater(<ExcelArgument(Description:="Temperature in Kelvin")> Temperature As Double) As Double

        Dim Thcon As Double

        Thcon = thconSatLiqTW(Temperature)
        Return Thcon

    End Function

    <ExcelFunction(Description:="Thermal conductivity of saturated steam at specified pressure, in W/m,K", Category:="IAPWS-IF97")>
    Public Function ThconSatVapPwater(<ExcelArgument(Description:="Pressure in bara")> Pressure As Double) As Double

        Dim Thcon As Double

        Thcon = thconSatVapPW(Pressure)
        Return Thcon

    End Function


    <ExcelFunction(Description:="Thermal conductivity of saturated steam at specified temperature, in W/m,K", Category:="IAPWS-IF97")>
    Public Function ThconSatVapTwater(<ExcelArgument(Description:="Temperature in Kelvin")> Temperature As Double) As Double

        Dim Thcon As Double

        Thcon = thconSatVapTW(Temperature)
        Return Thcon

    End Function

End Module
