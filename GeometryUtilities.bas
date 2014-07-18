'/***************************************************************************
'         VBA Utilities
'                             -------------------
'    begin                : 2013-07-23
'    copyright            : (C) 2013 by Walter Tsui
'    email                : waltertech426@gmail.com
' ***************************************************************************/

'/***************************************************************************
' *                                                                         *
' *   This program is free software; you can redistribute it and/or modify  *
' *   it under the terms of the GNU General Public License as published by  *
' *   the Free Software Foundation; either version 2 of the License, or     *
' *   (at your option) any later version.                                   *
' *                                                                         *
' ***************************************************************************/

Attribute VB_Name = "GeometryUtilities"
Option Compare Database

'Coordinate Transformation according to http://www.geodetic.gov.hk/data/pdf/explanatorynotes_c.pdf

'Coordinate Transform from HK1980 grid to WGS84 geographic in degree
Public Function CoorTransform_Hk1980ToWgs84(Easting, Northing, Optional Delimiter As String = "") As Variant
    'Initilalize Constant
    E0 = 836694.05
    N0 = 819069.8
    Lng0 = 114.178556
    Lat0 = 22.312133
    m_0 = 1
    M0 = 2468395.723
    a = 6378388
    e2 = 6.722670022 * (10 ^ (-3))
    
    LngLat_HK1980 = CoorTransform_GridToGeographic(E0, N0, Lng0, Lat0, m_0, M0, a, e2, Easting, Northing)

    Lng_WGS84 = LngLat_HK1980(0) + (8.8 / 3600)
    Lat_WGS84 = LngLat_HK1980(1) - (5.5 / 3600)
    
    
    If Delimiter = "" Then
        CoorTransform_Hk1980ToWgs84 = Array(Lng_WGS84, Lat_WGS84)
    Else
        CoorTransform_Hk1980ToWgs84 = Lng_WGS84 & Delimiter & Lat_WGS84
    End If
    
    
End Function


'Coordinate Transform from grid to geographic in degree
Public Function CoorTransform_GridToGeographic(E0, N0, Lng0, Lat0, m_0, M0, a, e2, Easting, Northing) As Variant
    'Meridian distance Coefficients
    A0 = 1 - (e2 / 4) - (3 * (e2 ^ 2) / 64)
    A2 = (3 / 8) * (e2 + ((e2 ^ 2) / 4))
    A4 = (15 / 256) * (e2 ^ 2)
    

    'Convert the Lat0 and Lng0 from degree to radian
    Lng0 = Lng0 * Pi / 180
    Lat0 = Lat0 * Pi / 180
    
    
    'Convert from grid to geographic
    'Calculate Lat_p by iteration of Meridian distance,
    E_Delta = Easting - E0
    N_delta = Northing - N0
    Mp = (N_delta + M0) / m_0
    
    Lat_min = -90 * Pi / 180
    Lat_max = 90 * Pi / 180

    Do While Abs(M - Mp) > 0.0000001
        Lat_p = (Lat_max + Lat_min) / 2
        M = a * (A0 * Lat_p - A2 * Sin(2 * Lat_p) + A4 * Sin(4 * Lat_p))
        
        If M >= Mp Then
            Lat_max = Lat_p
        Else
            Lat_min = Lat_p
        End If
        
    Loop
    
    
    t_p = Tan(Lat_p)
    v_p = a / ((1 - e2 * Sin(Lat_p) ^ 2) ^ (1 / 2))
    p_p = (a * (1 - e2)) / ((1 - e2 * Sin(Lat_p) ^ 2) ^ (3 / 2))
    W_p = v_p / p_p
    

    Lng = Lng0 + (1 / Cos(Lat_p)) * ((E_Delta / (m_0 * v_p)) - (1 / 6) * ((E_Delta / (m_0 * v_p)) ^ 3) * (W_p + 2 * (t_p ^ 2)))
    Lat = Lat_p - (t_p / ((m_0 * p_p))) * ((E_Delta ^ 2) / ((2 * m_0 * v_p)))


    CoorTransform_GridToGeographic = Array(Lng / Pi * 180, Lat / Pi * 180)
    
    
End Function


'Coordinate Transform from WGS84 geographic in degree to HK1980 grid
Public Function CoorTransform_Wgs84ToHK1980(Lng, Lat, Optional Delimiter As String = "") As Variant
    'Initilalize Constant
    E0 = 836694.05
    N0 = 819069.8
    Lng0 = 114.178556
    Lat0 = 22.312133
    m_0 = 1
    M0 = 2468395.723
    a = 6378388
    e2 = 6.722670022 * (10 ^ (-3))
    
    Lng_HK1980 = Lng - (8.8 / 3600)
    Lat_HK1980 = Lat + (5.5 / 3600)
    
    EastNorth_WGS84 = CoorTransform_GeographicToGrid(E0, N0, Lng0, Lat0, m_0, M0, a, e2, Lng_HK1980, Lat_HK1980)
    
    
    If Delimiter = "" Then
        CoorTransform_Wgs84ToHK1980 = Array(EastNorth_WGS84(0), EastNorth_WGS84(1))
    Else
        CoorTransform_Wgs84ToHK1980 = EastNorth_WGS84(0) & Delimiter & EastNorth_WGS84(1)
    End If
    
    
End Function


'Coordinate Transform from geographic in degree to grid
Public Function CoorTransform_GeographicToGrid(E0, N0, Lng0, Lat0, m_0, M0, a, e2, Lng, Lat) As Variant
    'Meridian distance Coefficients
    A0 = 1 - (e2 / 4) - (3 * (e2 ^ 2) / 64)
    A2 = (3 / 8) * (e2 + ((e2 ^ 2) / 4))
    A4 = (15 / 256) * (e2 ^ 2)
    

    'Convert Lat and Lng from degree to radian
    Lng0 = Lng0 * Pi / 180
    Lat0 = Lat0 * Pi / 180
    
    Lng = Lng * Pi / 180
    Lat = Lat * Pi / 180
    
    
    'Convert from geographic to grid
    Lng_Delta = Lng - Lng0
    M = a * (A0 * Lat - A2 * Sin(2 * Lat) + A4 * Sin(4 * Lat))

    t_s = Tan(Lat)
    v_s = a / ((1 - e2 * Sin(Lat) ^ 2) ^ (1 / 2))
    p_s = (a * (1 - e2)) / ((1 - e2 * Sin(Lat) ^ 2) ^ (3 / 2))
    W_s = v_s / p_s
    

    Easting = E0 + m_0 * v_s * (Lng_Delta * Cos(Lat) + (1 / 6) * (Lng_Delta ^ 3) * (Cos(Lat) ^ 3) * (W_s - t_s ^ 2))
    Northing = N0 + m_0 * ((M - M0) + v_s * ((Lng_Delta ^ 2) / 4) * Sin(2 * Lat))


    CoorTransform_GeographicToGrid = Array(Easting, Northing)
    
    
End Function
