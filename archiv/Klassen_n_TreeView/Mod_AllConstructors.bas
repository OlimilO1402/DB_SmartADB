Attribute VB_Name = "MNew"
Option Explicit

Public Function LandExplorer(Tvw As TreeView, StB As StatusBar, Optional aLand As Land) As LandExplorer
  Set LandExplorer = New LandExplorer: LandExplorer.New_ Tvw, StB, aLand
End Function

Public Function Land(LandName As String) As Land
  Set Land = New Land: Land.New_ LandName
End Function

Public Function City(CityName As String) As City
  Set City = New City: City.New_ CityName
End Function

Public Function Street(StreetName As String) As Street
  Set Street = New Street: Street.New_ StreetName
End Function

Public Function House(strHouseNumber As String) As House
  Set House = New House: House.New_ strHouseNumber
End Function

Public Function Family(FamilyName As String) As Family
  Set Family = New Family: Family.New_ FamilyName
End Function

Public Function Person(PersonName As String) As Person
  Set New_Person = New Person: Person.New_ PersonName
End Function

