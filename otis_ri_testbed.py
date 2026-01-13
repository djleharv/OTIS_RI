from lib import functions

print("Running otis_ri_testbed.py")

print("\tCreating Cargo Table")
functions.CreateCargoTable()

print("\tCreating Cargo PNMLs")
functions.CreateCargoPNMLs()

print("\tCreating Cargo LNGs")
functions.CreateCargoLNGs()

print("\tCreating Industry Files")
functions.CreateIndustries()

print("\tCreating Industry LNGs")
functions.CreateIndustryLNGs()

print("\tCreating Industry Help Text")
functions.CreateIndustryHelpText()

print("\tCreating Industry Help Texts LNGs")
functions.CreateIndustryHelpTextsLNGs()

print("\tCreating House PNMLs")
functions.CreateHousePNMLs()

print("\tCreating House LNGs")
functions.CreateHouseLNGs()

print("\tCreating Lang file")
functions.CreateLNGFile()
