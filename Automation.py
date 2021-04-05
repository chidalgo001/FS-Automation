import clr , ctypes
import os
import pymongo
import xml.etree.ElementTree as ET 
import sys
import re
import System
import json
import pystache
from xlrd import open_workbook
from System import TimeSpan , DateTime
clr.AddReference("Wtvision.Sports.Interfaces")
clr.AddReference("Wtvision.Link.Graphics.Base")
from Wtvision.Sports.Enums import EPlayerState, EPossession
from Wtvision.Link.Graphics import TagsViewBag 
from System.Threading import Thread
from Wtvision.Sports.Interfaces import IStats
from Wtvision.Sports.Constants import CollectionsNames
clr.AddReference("Wtvision.Sports.SportsCore")
from Wtvision.Sports import SportsCore
clr.AddReference("Wtvision.Sports.Models")
import Wtvision.Sports.Models
clr.ImportExtensions(Wtvision.Sports.Models)
clr.AddReference("Wtvision.Sports.Models.Football")
from Wtvision.Sports.Models.Football import FootballEvents
clr.AddReference("Wtvision.Sports.Interfaces")
from Wtvision.Sports.Interfaces import IStats

clr.AddReference("Microsoft.VisualBasic") 
from Microsoft.VisualBasic import Interaction , MsgBoxStyle



class _GraphicHandler:
	def __init__(self):
		self.Name = None
		self.ViewData = TagsViewBag();


def Execute(): 
	#Normal Bounds
	Globals["Bounds"] = [-0.1 , 2.05 , -2.5 , 0.15]
	Globals["Continue"] = False
	
	Core.ScriptRunner.Execute("MongoConnect", "Connect", Project.Name)
	
	if not Project.ProjectSettings["FootballEngineLocation"]:
		Project.ProjectSettings["FootballEngineLocation"] = "D:\\Generic_FootballCG\\"
	if not Project.ProjectSettings["GraphicsLocation"]:
		Project.ProjectSettings["GraphicsLocation"] = "R:\\Projects\\%s\\"%Project.Name
	#---------------	
	#Project.EventsManager.SubscribeEvent("IntelliflowController.UpdatingData", Script.Id, "OnUpdatingData", True)
	#if Globals["IntelliflowController.UpdatingData"]:
	#	IntelliflowController.UpdatingData -= Globals["IntelliflowController.UpdatingData"]
	#Globals["IntelliflowController.UpdatingData"] = OnUpdatingData
	#if Globals["IntelliflowController.GettingData"]:
	#	IntelliflowController.GettingData -= Globals["IntelliflowController.GettingData"]
	#Globals["IntelliflowController.GettingData"] = OnGettingData	
	#IntelliflowController.GettingData += Globals["IntelliflowController.GettingData"]
	#IntelliflowController.UpdatingData += Globals["IntelliflowController.UpdatingData"]
	#----------------	

	Project.EventsManager.UnsubscribeEvent("IntelliflowController.GettingData", Script.Id)
	Project.EventsManager.SubscribeEvent("IntelliflowController.BeforeProgram", Script.Id, "OnIntelliflowController_BeforeProgram", True)
	Project.EventsManager.SubscribeEvent("IntelliflowController.AfterProgram", Script.Id, "OnIntelliflowController_AfterProgram", True)
	Project.EventsManager.SubscribeEvent("IntelliflowController.AfterPreview", Script.Id, "OnIntelliflowController_AfterPreview", True)
	Project.EventsManager.SubscribeEvent("IntelliflowController.GettingData", Script.Id, "OnGettingData", False)
	Project.EventsManager.SubscribeEvent("GameOpened", Script.Id, "OnGameOpened", True)
	Project.EventsManager.SubscribeEvent("Game.SelectedTeamChanged", Script.Id, "OnGame_SelectedTeamChanged", True)
	Project.EventsManager.SubscribeEvent("Game.SelectedPlayerChanged", Script.Id, "OnGame_SelectedPlayerChanged", True)
	Project.EventsManager.SubscribeEvent("Game.SelectedStatsChanged", Script.Id, "OnGame_SelectedStatsChanged", True)
	Project.EventsManager.SubscribeEvent("GameOnline.ScoreChanged", Script.Id, "OnlineChanged", True)
	Project.EventsManager.SubscribeEvent("GameOnline.ClockTimeChanged", Script.Id, "OnlineChanged", True)
	Project.EventsManager.SubscribeEvent("GameClosed", Script.Id, "OnGameClosed", True)
	
	#UnsubscribeAll()
	
	
def UnsubscribeAll():

	print "Unsubscribing to all registered events"
	Project.EventsManager.UnsubscribeEvent("IntelliflowController.GettingData", Script.Id)
	Project.EventsManager.UnsubscribeEvent("IntelliflowController.BeforeProgram", Script.Id)
	Project.EventsManager.UnsubscribeEvent("IntelliflowController.AfterProgram", Script.Id)
	Project.EventsManager.UnsubscribeEvent("IntelliflowController.AfterPreview", Script.Id)
	Project.EventsManager.UnsubscribeEvent("IntelliflowController.GettingData", Script.Id)
	Project.EventsManager.UnsubscribeEvent("GameOpened", Script.Id)
	Project.EventsManager.UnsubscribeEvent("GameOnline.ScoreChanged", Script.Id)
	Project.EventsManager.UnsubscribeEvent("GameOnline.ClockTimeChanged", Script.Id)
	Project.EventsManager.UnsubscribeEvent("GameClosed", Script.Id)


def OnlineChanged(sender, args):
	
	try:		
		for graphic in Globals["LiveUpdates"]:
			if IntelliflowController.IsGraphicOnAir(graphic, OutputManager.ActiveChannelOutput):
				"""
				file = "%s\\%s.footballcfg"%(Project.ProjectSettings["FootballEngineLocation"],	graphic.Name)
				if os.path.exists(file):
					gfx = _GraphicHandler()
					gfx.Name = graphic.Name
					ParseXML(file, gfx, True)
				"""
				mongoItem = graphic.Name
		
				found , config = CheckMongoItem(mongoItem)
					
				if found:
					gfx = _GraphicHandler()
					gfx.Name = graphic.Name
					ParseMongo(mongoItem, gfx , config, True)
				
					onAirGraphics = IntelliflowController.GetGraphicOnAirItems(gfx.Name,  OutputManager.ActiveChannelOutput)
					
					IntelliflowController.FillData(onAirGraphics, gfx.ViewData, True)	
					
	except Exception as e:
		print "Error in OnlineChanged: %s"%e

def CleanTags(graphic):
	
	for tag in graphic.Scene.Tags:
		tag.Value = " "
	print "Graphic Cleaned"
	return

#-- Will check if the name of the graphic is already stored i nthe DB and if it has any configurations saved along with it
def CheckMongoItem(name):
	itemName = ''
	config = ''
	for row in Globals["mydb"].Graphics.find( {"name":name} ):	
		itemName = row["name"]
		config = row["config"]
		
	if itemName and config:
		return True , config
		
	else:
		if not itemName:
			print "No Document Named '%s' Found on DB"%name
		elif not config:
			print "Document Named '%s' Has No Config Found on DB"%name 
			
	return False , False


#-- Returns the numerical values of a selected range ( start_col, start_row, end_col, end_row )		
def get_cell_range(start_col, start_row, end_col, end_row, sheet):
		
    result = [sheet.row_values(row, start_colx=start_col, end_colx=end_col+1) for row in xrange(start_row, end_row+1)]
    #print "%s,%s,%s,%s"%(start_col, start_row, end_col, end_row)

    return result 
  
#-- Returns the Numerical value of a cell as a tuple, i.e. A2 = (0 , 1) , B1 = (1 , 0)
def getCellIndex(cellReference):
	
	result = None,None
	
	text = re.findall(r"[A-Z]+",cellReference)			
	numbers = re.findall(r"[0-9]+",cellReference)
	#print "%s=%s,%s"%(cellReference,text,numbers)
	
	if len(text) and len(numbers):
		return int(ord(text[0])-65), (int(numbers[0])-1)
		
	return result

def OnGettingData(sender, args):
	
	try:
		graphic = args.Graphics[0].Graphic
		
		#file = "%s\\%s.footballcfg"%(Project.ProjectSettings["FootballEngineLocation"],	graphic.Name)
		
		mongoItem = graphic.Name
		print "Getting Mongo Item: %s"%mongoItem

		found , config = CheckMongoItem(mongoItem)
		
		if found:
			Globals["graphic"] = args.Graphics[0]
			args.Cancel = ParseMongo(mongoItem, graphic , config)
		
			if graphic.LiveUpdate:
				if not Globals["LiveUpdates"]:
					Globals["LiveUpdates"] = []
				if not graphic in Globals["LiveUpdates"]:
					Globals["LiveUpdates"].append(graphic)


		else:
			pass
	
	except Exception as e:
		print e	
	

def ParseMongo(file, graphic, config, forUpdate=False):

	PenaltiesForm = "PenaltiesForm"
	Globals["Continue"] = False
	
	if not graphic:
		return True;	
	
	GraphicsLocation = Project.ProjectSettings["GraphicsLocation"] 	
	CompetitionLogo = Dictionary.Translate("CompetitionLogo")

	xml = ET.fromstring(config);
			
	for xmlNode in xml.getchildren():
	
		#Handling of XL inputs
		if xmlNode.tag == "XLS":
			dic = {}
			workbook = None
			
			file = Project.ProjectSettings["FootballEngineLocation"] + xmlNode.attrib["Source"]
			
			if os.path.exists(file):
			
				with open_workbook(file) as workbook:
					page = workbook.sheet_by_name(xmlNode.attrib["SheetName"])
					
					for child in xmlNode:
						columnValue,rowValue = getCellIndex(child.text)
						
						if child.tag == 'Cell':
							exec( "%s ='%s'"%(child.attrib['Name'],page.cell(rowValue,columnValue).value) )
							
						elif child.tag == 'Range':
							
							vector = child.text.split(":")
							
							if len(vector) == 2:
								startCell = vector[0]
								endCell = vector[1]
								start_col, start_row =  getCellIndex(startCell)
								end_col, end_row = getCellIndex(endCell)
								data = get_cell_range(start_col, start_row, end_col, end_row, page)
								
								try:
									exec ("%s = data"%(child.attrib['Name']))
								except Exception as e:						
									print "Error: %s"%e
									
					workbook.release_resources()
				workbook = None

		if xmlNode.tag == "Validation":

			for child in xmlNode:		
				dic = {}
				results = pystache.parse(child.attrib["When"])
				m = re.findall(r"key='(.*?)'", "%s"%results)
				
				for i in range(0, len(m)):
					try:
						exec ("dic['"+m[i]+"']="+m[i].replace("|","."))
					except Exception as e:
						print "Error in evaluating: %s"%e
						return True
						
				validation = pystache.render(child.attrib["When"], dic)
				
				
				if (validation) and (validation=='True'):
					return True
				
		#Linear Tags
		if xmlNode.tag == "ExportList":
			
			if "ClearTags" in xmlNode.attrib and xmlNode.attrib["ClearTags"]:
				CleanTags(Globals["graphic"])
			
			for child in xmlNode:	
				
				dic = {}
				if forUpdate and (not "Update" in child.attrib):					
					continue;
				
				results = pystache.parse(child.attrib["Value"])
				
				m = re.findall(r"key='(.*?)'", "%s"%results)
				
				for i in range(0, len(m)):
					try:
						
						exec ("dic['"+m[i]+"']="+m[i].replace("|","."))
						
					except Exception as e:
						print "Error in evaluating: %s"%e
						print "dic['"+m[i]+"']="+m[i].replace("|",".")
				
				if child.attrib["Type"].Contains("Color"):	
					color = pystache.render(child.attrib["Value"], dic)
					graphic.ViewData.SetString(child.attrib["Name"], GetColorRGB(color))
					
				elif child.attrib["Type"].Contains("Date"):
					date = pystache.render(child.attrib["Value"], dic)
					date = FormatDateTime(date)
					graphic.ViewData.SetString(child.attrib["Name"], date.upper())
					
				elif child.attrib["Type"].Contains("Countdown"):
					
					graphic.ViewData.SetString(child.attrib["Name"], SetCountdown() )
					
				elif "Translate" in child.attrib and child.attrib["Translate"]:
					graphic.ViewData.SetString(child.attrib["Name"],Dictionary.Translate(pystache.render(child.attrib["Value"], dic)).upper() )
		
				else:
					#return		
					#print "%s - %s"%(child.attrib["Name"], pystache.render(child.attrib["Value"], dic))
					graphic.ViewData.SetString(child.attrib["Name"], pystache.render(child.attrib["Value"], dic).upper())
				
				#print "%s - %s"%(dic , (child.attrib["Name"]))	
			
		#Handling lists
		if xmlNode.tag == "Multiplex":
		
			dic = {}
			try:
				if not xmlNode.attrib['Collection'] in locals():
				
					exec "dic[xmlNode.attrib['Name']]=%s()"%(xmlNode.attrib["Collection"])
				else:		
					exec "dic[xmlNode.attrib['Name']]=%s"%(xmlNode.attrib["Collection"])
				
			except Exception as e:
				print "ERROR in List Root: %s"%e
				print "dic[xmlNode.attrib['Name']]=%s()"%(xmlNode.attrib["Collection"])
				
			min = int(xmlNode.attrib["MinValue"])
			max = int(xmlNode.attrib["MaxValue"])
			
			startat = 0
			try:
				exec "%s=dic[xmlNode.attrib['Name']]"%(xmlNode.attrib['Name'])
				print "%s=dic[xmlNode.attrib['Name']]"%(xmlNode.attrib['Name'])
			except Exception as e:
				print e
				print "%s=dic[xmlNode.attrib['Name']]"%(xmlNode.attrib['Name'])
				
			
			# Sorting 
			if "Sort" in xmlNode.attrib and xmlNode.attrib["Sort"]:
				try:
					exec "%s"%(xmlNode.attrib['Sort'])
					exec "dic[xmlNode.attrib['Name']]=%s"%xmlNode.attrib['Name']
					
					#max = len(dic[xmlNode.attrib['Name']])
				except Exception as e:
					print e
					print "%s"%(xmlNode.attrib['Sort'])
					
			#Formation - This is to get the name of the formation
			if "Formation" in xmlNode.attrib and xmlNode.attrib["Formation"]:
				try:	
					print "dic[xmlNode.attrib['Formation']]=%s()"%("GetFormationName")
					exec "dic[xmlNode.attrib['Formation']]=%s()"%("GetFormationName")
					
				except Exception as e:
					print e
					print "%s"%(xmlNode.attrib['Formation'])
					
			#Continue		
			if "Continue" in xmlNode.attrib and xmlNode.attrib["Continue"] == "True":
				Globals["Continue"] = True
			
			max = len(dic[xmlNode.attrib['Name']]) 	#This will make sure that the number of iterations is the same as the number if items to show.
													#This is done to avoid repeated items on the graphic
			
			
			#this is done to let execute the len(items) = 0 when there is an "[]" as a result,
			#otherwise, the tag will not enter the loop below.
			#------------------------
			if max == 0:
				max = 1
			#------------------------	
			
			#Multiplex Items BEGIN
			for i in range(min,max+1):
				
				tagNumbering = str(i).zfill(len(xmlNode.attrib["MinValue"]))
				
				dic[xmlNode.attrib['IteratorName']]=dic[xmlNode.attrib['Name']][startat] if len(dic[xmlNode.attrib['Name']]) > startat else None
	
				try:	
					#print "%s=dic[xmlNode.attrib['IteratorName']]"%( xmlNode.attrib['IteratorName'])
					exec "%s=dic[xmlNode.attrib['IteratorName']]"%( xmlNode.attrib['IteratorName'])
					
				except Exception as e:
					print e
					print "%s=dic[xmlNode.attrib['IteratorName']]"%( xmlNode.attrib['IteratorName'])
				
				startat = startat + 1
				#Keys in the Expression Begin
				for child in xmlNode:	
	
					results = pystache.parse(child.attrib["Value"])	
					if forUpdate and (not "Update" in child.attrib):
						continue;
						
					m = re.findall(r"key='(.*?)'", "%s"%results)
					
					#Key Elements to get Values for each one
					for i in range(0, len(m)):		
	
						try:						
							if dic[xmlNode.attrib['IteratorName']]:
								
								if not m[i] == "tactic":
									exec ("dic['"+m[i]+"']="+m[i].replace("|","."))
									
						except Exception as e: 
							print str(e) + " In Multiplex"
							if dic[xmlNode.attrib['IteratorName']]:
								print ("dic['"+m[i]+"']="+m[i].replace("|","."))
							else:
								print ("dic['"+m[i]+"']=None")
		
					#Key Elements end
					if child.attrib["Type"].Contains("Color"):	
						color = pystache.render(child.attrib["Value"], dic)
						graphic.ViewData.SetString(child.attrib["Name"], GetColorRGB(color))
						
					elif "Translate" in child.attrib and child.attrib["Translate"]:
						graphic.ViewData.SetString(child.attrib["Name"].replace("#",tagNumbering),Dictionary.Translate(pystache.render(child.attrib["Value"], dic)).upper() )
						
					else:
						graphic.ViewData.SetString(child.attrib["Name"].replace("#",tagNumbering), pystache.render(child.attrib["Value"], dic).upper())
					
					#print  "%s -- %s"%(child.attrib["Name"].replace("#",tagNumbering) , pystache.render(child.attrib["Value"], dic) )
					#Keys in the Expression END
					
				#Multiplex Items END
		
		# Handles Sponsors
		if xmlNode.tag == "Sponsor":

			for child in xmlNode:
				dic = {}	

				results = pystache.parse(child.attrib["Value"])
				
				m = re.findall(r"key='(.*?)'", "%s"%results)
				
				for i in range(0, len(m)):
					try:
						
						exec ("dic['"+m[i]+"']="+m[i].replace("|","."))
						
					except Exception as e:
						print "Error in evaluating: %s"%e
						print "dic['"+m[i]+"']="+m[i].replace("|",".")

				img = GetSponsor(graphic.Name) 
			
				if child.attrib["Type"].Contains("Visibility"):
					if img == None or img == "" :
						child.attrib["Value"] = "1"
				
					else:
						child.attrib["Value"] = "2"
	
					graphic.ViewData.SetString(child.attrib["Name"], child.attrib["Value"])
				
				elif child.attrib["Type"].Contains("Image"):
					graphic.ViewData.SetString(child.attrib["Name"], pystache.render(child.attrib["Value"], dic))
				#print "%s - %s" %( child.attrib["Name"], pystache.render(child.attrib["Value"], dic) )
					
	return False
#=================================================	
#------------- Auxiliary functions ---------------

def GetSponsor(name):
	itemName = ''
	sponsor = ''
	for row in Globals["mydb"].Graphics.find( {"name":name} ):	
		itemName = row["name"]
		sponsor = row["sponsor"]
	if itemName and sponsor:
		return sponsor

def FormatDateTime(date):
	
	localDateTime = DateTime.Parse(date)
	return "%s %sh%s"%(localDateTime.ToString("MMM dd"), localDateTime.ToString("HH"),localDateTime.ToString("mm")  )

	
def FormattedTime(clock):
	try:
		regularclocks = {"SS":"ss",
		"SSF":"ss\.f",
		"SSFF":"ss\.ff", 
		"SSFFF":"ss\.fff", 
		"MMSS":"mm\:ss", 
		"MMSSF":"mm\:ss\.f",
		"MMSSFF":"mm\:ss\.ff",
		"MMSSFFF":"mm\:ss\.fff",
		"HHMMSS":"hh\:mm\:ss",
		"HHMMSSF": "hh\:mm\:ss\.f",
		"HHMMSSFF":"hh\:mm\:ss\.ff", 
		"HHMMSSFFF":"hh\:mm\:ss\.fff"}
		if GameOnline.Clock.Mask.ToString() in regularclocks.keys():
			return clock.ToString(regularclocks[GameOnline.Clock.Mask.ToString()])
		elif GameOnline.Clock.Mask.ToString()=="MMMSS":
			minutes = str(int(GameOnline.Clock.ElapsedTime.TotalSeconds / 60))
			if len(minutes)<2:
				minutes= minutes.zfill(2)
			seconds = str(int(GameOnline.Clock.ElapsedTime.TotalSeconds % 60))
			if len(seconds)<2:
				seconds= seconds.zfill(2)
				
			return "%s:%s"%(minutes,seconds)
		return clock.ToString()
	except:
		return ""

#================== Tactical Functions ==================

def FillTactics():
	
	ResetBounds()

	return getTacticPlayers()

def FillHomeTacticsAR():

	#Bounds for the Lineup_AR graphic 
	Tactics.FieldBounds.XMax = 32
	Tactics.FieldBounds.XMin = -32
	Tactics.FieldBounds.YMax = 49
	Tactics.FieldBounds.YMin = -49
	
	return getTacticPlayers("Home")
	
def FillAwayTacticsAR():

	#Bounds for the Lineup_AR graphic 
	Tactics.FieldBounds.XMax = 32
	Tactics.FieldBounds.XMin = -32
	Tactics.FieldBounds.YMax = 49
	Tactics.FieldBounds.YMin = -49
	
	return getTacticPlayers("Away")


#Used for graphics tactical graphics that require a different bound from the main formation graphic
def FillTacticsVideos():
	#CleanTags(Globals["OnAir"])
	#Bounds for the Lineup_Video graphic 
	Tactics.FieldBounds.XMax = 1.6
	Tactics.FieldBounds.XMin = -0.1
	Tactics.FieldBounds.YMax = 0.05
	Tactics.FieldBounds.YMin = -1.5
	
	return getTacticPlayers()

def getTacticPlayers(team = None):

	TacticPlayers= []
	
	if team == "Home":
		TacticPlayers = list(Tactics.HomeTeamConfig.TacticPlayers)
	elif team == "Away":
		TacticPlayers = list(Tactics.AwayTeamConfig.TacticPlayers)
	elif Game.SelectedTeam.IsHomeTeam:
		TacticPlayers = list(Tactics.HomeTeamConfig.TacticPlayers)
	else :
		TacticPlayers = list(Tactics.AwayTeamConfig.TacticPlayers)
	

	TacticPlayers = sorted(TacticPlayers, key=lambda p: (100*p.Y)+p.X)
	
	return TacticPlayers

# Returns the name of the formation for the selected team
def GetFormationName():

	Globals["Lineup"] = []
	
	if Game.SelectedTeam.IsHomeTeam:
		TacticName = Tactics.HomeTeamConfig.Name
	else:
		TacticName = Tactics.AwayTeamConfig.Name
		
	#Globals["Lineup"].extend( TacticName.split("-") )
	Globals["Lineup"] = TacticName.split("-")
	return TacticName

# sorting order for the tactical lineups
def TacticsCompY(obj1, obj2):
	#print "%s < %s"%(obj1.GraphicPosition.Y , obj2.GraphicPosition.Y)
	return obj1.GraphicPosition.Y < obj2.GraphicPosition.Y

#========================================================================

def SelectedTeamInPlay():
	SelectedTeamInPlay = []
	for a  in Game.SelectedTeam.Players:
		if ("%s"%a.State) == 'InPlay':
			SelectedTeamInPlay.append(a)
	return SelectedTeamInPlay
	
	
#==================== BENCH FUNCTIONS =======================
def SelectedTeamInBench():

	SelectedTeamInBench = []
	for a  in Game.SelectedTeam.Players:
		if ("%s"%a.State) == 'InBench':
			SelectedTeamInBench.append(a)
	return SelectedTeamInBench

def HomeDoubleBench():
	return DoubleBench("home")
	
def AwayDoubleBench():
	return DoubleBench("away")

# Use DoubleBench when both benches need to have the same amount of lines (share vlines)
def DoubleBench():
	Home = Bench("home")
	Away = Bench("away")
	
	if len(Away)>len(Home):
		for a in range(len(Home), len(Away)):
			Home.append(None)
	elif len(Home)>len(Away):
		for a in range(len(Away), len(Home)):
			Away.append(None)
			

	return zip(Home,Away)
	
# Use the individual Bench function when the graphic has independent vlines 	
def HomeBench():
	home = Bench("home")
	return home

def AwayBench():
	away = Bench("away")
	return away

# Funtion that returns the appropiate bench sorted
def Bench(team):
	TeamInBench = []
	
	if team.upper() == "HOME":		
		for a  in Game.HomeTeam.Players:
			if ("%s"%a.State) == 'InBench':
				TeamInBench.append(a)
		return sorted(TeamInBench, key=lambda x: x.IsGoalKeeper*-1000+x.Number, reverse=False)
	if team.upper() == "AWAY":		
		for a  in Game.AwayTeam.Players:
			if ("%s"%a.State) == 'InBench':
				TeamInBench.append(a)
		return sorted(TeamInBench, key=lambda x: x.IsGoalKeeper*-1000+x.Number, reverse=False)

#===============================================================	

def FetchPlayerID(id, team):
	if team == "Home":
		team = Game.HomeTeam
	elif team == "Away":
		team = Game.AwayTeam

	for player in team.Players:
		if str(player.Id) == id:
			
			return player
	return
	
def GetHomeScorers():
		
	return  Globals["Scorers"]["Home"]

def GetAwayScorers():
		
	return  Globals["Scorers"]["Away"]

#-- takes in either Hex or RGB color format, and converts to Usable RGB to use in R3
def GetColorRGB(color):
	try:		
		if 'str' in str(type(color)):
		
			if color.count(',') == 2: # RGB to Color
				rgb = map(int, color.split(','))
				rgb = str(tuple(rgb))
				return rgb[1:-1:]
				
			if '#' in color: # HEX to Color
			
				color = color.ljust(9,"F")
				hex = color.lstrip('#')
				rgb = str(tuple(int(hex[i:i+2], 16) for i in (2, 4,6)) )
				return rgb[1:-1:]
				
		return ''
	except:
		print "Error in converting Hex To RGB"
		return ''
	
def Referees():
	return list(Game.Referees)		
	
def CollectedStats():	
	return list(Game.SelectedStats)

def InPlayToOut():
	player = Game.SelectedPlayer
	player.State = EPlayerState.SubstOut
	player.SubstituteIn = False
	player.SubstituteOut = True
	
def BenchToInPlay():
	player = Game.SelectedPlayer
	player.State = EPlayerState.InPlay
	player.SubstituteIn = True
	player.SubstituteOut = False
	
def SetPlayerGameStat():
	player = Game.SelectedPlayer
	player.Stats.Shots += 1
	
def SetTeamGameStat():
	team = Game.SelectedTeam
	team.Stats.Shots += 1
	
	
def SetPlayerCompStat():
	player = Game.SelectedPlayer
	player.CompetitionStats.Goals += 1
	
def SetTeamCompStat():
	team = Game.SelectedTeam
	team.CompetitionStats.Goals += 1

def CalcHPenaltyScore():
	team="Home"
	#print Globals["Penalties"][team]
	lst = ["0" , "0" , "0" ,"0" ,"0"]


	for attempt in Globals["Penalties"][team]:
	
		position = int( attempt[-1:] ) -1
		
		if attempt.Contains("Converted"):
			lst[ position ] = "1"

		else:
			lst[ position ] = "2"

			

	return lst

def CalcAPenaltyScore():
	team="Away"
	#print Globals["Penalties"][team]
	lst = ["0" , "0" , "0" ,"0" ,"0"]


	for attempt in Globals["Penalties"][team]:
	
		position = int( attempt[-1:] ) -1
		
		if attempt.Contains("Converted"):
			lst[ position ] = "1"
			
		else:
			lst[ position ] = "2"

	return lst


def SetCountdown():
	
	value = Interaction.InputBox("How long do you want the timer for?\nFormat: \t mm:ss","CountDown Timer","00:00") 
	return value

#========================================

def OnGame_SelectedTeamChanged(sender, args):
	Globals["OperationForm"].Team = Game.SelectedTeam

def OnGame_SelectedPlayerChanged(sender, args):
	Globals["OperationForm"].Player = Game.SelectedPlayer

def OnGame_SelectedStatsChanged(sender, args):
	if len(list(Game.SelectedStats))==1:
		Globals["OperationForm"].Stats = list(Game.SelectedStats)[0]


#========================================

def OnGameOpened(sender, args):

	Globals["Scorers"] = { "Home": [] , "Away": [] }	
	Globals["SportsCore"] = SportsCore.Instance
	
	try:
		LoadTeamStats()
	except Exception as e:
		print e		
	try:
		LoadPlayerStats()
	except Exception as e:
		print e
	try:
		path = ("%s/Saves/%s_Scorers.json")%( Project.ProjectSettings["FootballEngineLocation"], str(Game.Id) )
		if os.path.exists(path):
			with open(path,"r") as json_file:		
				Globals["Scorers"] = json.loads( json_file.read() )
		else:
			Globals["Scorers"] = { "Home": [] , "Away": [] }	
		 
	except Exception as e:
		print e

def ResetBounds():

	# Global["Bounds"] = [XMin , XMax , YMin , YMax]
	Tactics.FieldBounds.XMax = Globals["Bounds"][1]
	Tactics.FieldBounds.XMin = Globals["Bounds"][0]
	Tactics.FieldBounds.YMax = Globals["Bounds"][3]
	Tactics.FieldBounds.YMin = Globals["Bounds"][2]

def OnGameClosed(sender, args):

	ResetBounds()	
	SaveTeamStats()
	SavePlayerStats()
	UnsubscribeAll()


def SaveTeamStats():

	SaveStats(Game.HomeTeam.Stats)
	SaveStats(Game.AwayTeam.Stats)
	
def SavePlayerStats():
	homePlayers = Game.HomeTeam.Players
	awayPlayers = Game.AwayTeam.Players
	
	for player in homePlayers:
		SaveStats(player.Stats)
		
	for player in awayPlayers:
		SaveStats(player.Stats)


def SaveStats(statsObj):
	try:
		Game.Storage.Replace[IStats](CollectionsNames.Stats, statsObj.Id, statsObj, True)
	except Exception as e:
		print e

def LoadTeamStats():

	hometeam = Game.HomeTeam
	awayteam = Game.AwayTeam
	
	hometeamStats = Globals["SportsCore"].GetTeamGameStats(hometeam.Id)
	awayteamStats = Globals["SportsCore"].GetTeamGameStats(awayteam.Id)
	
	temp = dir(Game.HomeTeam.Stats)
	for stat in temp:
		if not stat.Contains("__"):

			try:
				exec("hometeam.Stats.%s = hometeamStats.%s")%(stat,stat)
				exec("awayteam.Stats.%s = awayteamStats.%s")%(stat,stat)
			except:
				pass
	
def LoadPlayerStats():
	homePlayers = Game.HomeTeam.Players
	awayPlayers = Game.AwayTeam.Players
	
	temp = dir(Game.HomeTeam.Stats)
	
	for player in homePlayers:		
		playerStats = Globals["SportsCore"].GetPlayerGameStats(player.Id)
	
		for stat in temp:
			if not stat.Contains("__"):
				try:
					exec("player.Stats.%s = playerStats.%s")%(stat,stat)
				except:
					pass
	
	for player in awayPlayers:		
		playerStats = Globals["SportsCore"].GetPlayerGameStats(player.Id)
	
		for stat in temp:
			if not stat.Contains("__"):
				try:
					exec("player.Stats.%s = playerStats.%s")%(stat,stat)
				except:
					pass
				

def OnIntelliflowController_AfterProgram(sender, args):

	gfx = args.Graphics[0].Graphic.Name.ToString()

	if gfx == "Substitution" or gfx == "PermanentClock":	
		Core.ScriptRunner.Execute("After_Program_Handler", "PermanentClock", Project.Name, args)	

	if gfx == "Substitution" or gfx == "Substitution-HT":	
		Core.ScriptRunner.Execute("After_Program_Handler", "Substitution", Project.Name, args)			

	if gfx == "LineUpPhoto":	
		Core.ScriptRunner.Execute("After_Program_Handler", "LineUpPhoto", Project.Name, args)


def OnIntelliflowController_BeforeProgram(sender, args):
	#print Globals["Continue"]
	if Globals["Continue"] == False:
		return
	"""	
	for i in range(1,12):
		strindx = str(i).zfill(2)
		args.Graphics[0].Scene.Tags["vRGBTacticNumber%s"%(strindx)].Value = "-6316129"
		
	args.Graphics[0].Scene.Tags["vRGBTacticNumber01"].Value = "-631"
	"""




def OnIntelliflowController_AfterPreview(sender, args):
	pass
	"""
	#IntelliflowController.TakeIn(args.Graphics[0].Graphic, OutputManager.ActiveChannelOutput, IntelliflowController.EWorkflowType.Preview)	
	IntelliflowController.Start(args.Graphics, IntelliflowController.EWorkflowType.Preview)
	args.Graphics[0].Scene.GoToFrame("Change_In", 89)
	"""
