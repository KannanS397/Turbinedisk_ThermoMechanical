# -*- coding: mbcs -*-
from part import *
from material import *
from section import *
from assembly import *
from step import *
from interaction import *
from load import *
from mesh import *
from optimization import *
from job import *
from sketch import *
from visualization import *
from connectorBehavior import *
import os
import xlrd
import xlsxwriter
wb = xlrd.open_workbook("TD_LHS_100.xlsx")
sheet = wb.sheet_by_index(0)
# sheet.cell_value(:, :)
#Design variable
P15 = sheet.col_values(0)
P17 = sheet.col_values(1)
P18 = sheet.col_values(2)
P19 = sheet.col_values(3)
P20 = sheet.col_values(4)
P21 = sheet.col_values(5)

#############################
#Load variable
W = sheet.col_values(6)
F = sheet.col_values(7)
E = sheet.col_values(8)
v = sheet.col_values(9)
k = sheet.col_values(10)
a = sheet.col_values(11)
Tt = sheet.col_values(12)
Tb = sheet.col_values(13)
Ht = sheet.col_values(14)
Hb = sheet.col_values(15)
#############################
P16 = 21
workbook = xlsxwriter.Workbook('TD_Coupled.xlsx')
wso = workbook.add_worksheet()
#############################
for i in range(len(P15)):
    print 'Model :', i+1

    mdb.Model(name='Model-%s'%(i+1), modelType=STANDARD_EXPLICIT)
    session.viewports['Viewport: 1'].setValues(displayedObject=None)
    
    mdb.openStep(
        'C:/Kannan/Suja/Turbine_Disk_New_V1 v3.step'
        , scaleFromFile=OFF)
    mdb.models['Model-%s'%(i+1)].ConstrainedSketchFromGeometryFile(geometryFile=mdb.acis, 
        name='Turbine_Disk_New_V1 v3')
    mdb.models['Model-%s'%(i+1)].ConstrainedSketch(name='__profile__', sheetSize=1000.0)
    mdb.models['Model-%s'%(i+1)].sketches['__profile__'].sketchOptions.setValues(
        viewStyle=AXISYM)
    mdb.models['Model-%s'%(i+1)].sketches['__profile__'].ConstructionLine(point1=(0.0, 
        -500.0), point2=(0.0, 500.0))
    mdb.models['Model-%s'%(i+1)].sketches['__profile__'].FixedConstraint(entity=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[2])
    mdb.models['Model-%s'%(i+1)].sketches['__profile__'].retrieveSketch(sketch=
        mdb.models['Model-%s'%(i+1)].sketches['Turbine_Disk_New_V1 v3'])
    mdb.models['Model-%s'%(i+1)].sketches['__profile__'].move(objectList=(
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[5], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[6], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[7], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[8], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[9], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[10], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[11], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[12], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[13], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[14], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[15], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[16], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[17], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[18], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[19], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[20], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[21], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[22], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[23], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[24], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[25], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[26], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[27], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[28], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[29], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[30], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[31], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[32], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[33], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[34], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[35], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[36], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[37], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[38]), vector=(120.0, 
        0.0))
    mdb.models['Model-%s'%(i+1)].sketches['__profile__'].Line(point1=(120.0, 0.0), point2=
        (820.0, 0.0))
    mdb.models['Model-%s'%(i+1)].sketches['__profile__'].HorizontalConstraint(
        addUndoState=False, entity=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[39])
    mdb.models['Model-%s'%(i+1)].sketches['__profile__'].PerpendicularConstraint(
        addUndoState=False, entity1=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[25], entity2=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[39])
    mdb.models['Model-%s'%(i+1)].sketches['__profile__'].CoincidentConstraint(
        addUndoState=False, entity1=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].vertices[42], entity2=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[25])
    mdb.models['Model-%s'%(i+1)].sketches['__profile__'].EqualDistanceConstraint(
        addUndoState=False, entity1=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].vertices[24], entity2=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].vertices[25], midpoint=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].vertices[42])
    mdb.models['Model-%s'%(i+1)].sketches['__profile__'].CoincidentConstraint(
        addUndoState=False, entity1=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].vertices[43], entity2=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[8])
    mdb.models['Model-%s'%(i+1)].sketches['__profile__'].EqualDistanceConstraint(
        addUndoState=False, entity1=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].vertices[3], entity2=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].vertices[4], midpoint=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].vertices[43])
    mdb.models['Model-%s'%(i+1)].sketches['__profile__'].setAsConstruction(objectList=(
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[39], ))
    mdb.models['Model-%s'%(i+1)].sketches['__profile__'].delete(objectList=(
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[8], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[9], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[10], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[11], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[12], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[13], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[14], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[15], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[16], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[17], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[18], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[19], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[20], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[21], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[22], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[23], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[24], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[25], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].constraints[81], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].constraints[84], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].constraints[86]))
    mdb.models['Model-%s'%(i+1)].sketches['__profile__'].Line(point1=(120.0, -151.0), 
        point2=(120.0, 0.0))
    mdb.models['Model-%s'%(i+1)].sketches['__profile__'].VerticalConstraint(addUndoState=
        False, entity=mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[40])
    mdb.models['Model-%s'%(i+1)].sketches['__profile__'].PerpendicularConstraint(
        addUndoState=False, entity1=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[26], entity2=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[40])
    mdb.models['Model-%s'%(i+1)].sketches['__profile__'].Line(point1=(820.0, 0.0), point2=
        (820.0, -28.5022161495221))
    mdb.models['Model-%s'%(i+1)].sketches['__profile__'].VerticalConstraint(addUndoState=
        False, entity=mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[41])
    mdb.models['Model-%s'%(i+1)].sketches['__profile__'].PerpendicularConstraint(
        addUndoState=False, entity1=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[39], entity2=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[41])
    mdb.models['Model-%s'%(i+1)].sketches['__profile__'].DistanceDimension(entity1=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[2], entity2=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[40], textPoint=(
        54.9898376464844, -73.4044570922852), value=100.0)
    mdb.models['Model-%s'%(i+1)].sketches['__profile__'].ObliqueDimension(textPoint=(
        159.975830078125, -190.588333129883), value=P15[i], vertex1=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].vertices[25], vertex2=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].vertices[26])
    mdb.models['Model-%s'%(i+1)].sketches['__profile__'].DistanceDimension(entity1=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[29], entity2=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[34], textPoint=(
        309.055938720703, -63.529411315918), value=P16)
    mdb.models['Model-%s'%(i+1)].sketches['__profile__'].RadialDimension(curve=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[37], radius=P17[i], 
        textPoint=(770.29443359375, 26.0043411254883))
    mdb.models['Model-%s'%(i+1)].sketches['__profile__'].RadialDimension(curve=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[33], radius=P18[i], 
        textPoint=(371.34765625, -169.521575927734))
    mdb.models['Model-%s'%(i+1)].sketches['__profile__'].RadialDimension(curve=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[28], radius=P19[i], 
        textPoint=(262.862091064453, -74.0627975463867))
    mdb.models['Model-%s'%(i+1)].sketches['__profile__'].ObliqueDimension(textPoint=(
        59.1892700195313, -89.2045364379883), value=P20[i], vertex1=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].vertices[25], vertex2=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].vertices[42])
    mdb.models['Model-%s'%(i+1)].sketches['__profile__'].DistanceDimension(entity1=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].vertices[39], entity2=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[39], textPoint=(
        648.510620117188, -19.4208602905273), value=P21[i])
    mdb.models['Model-%s'%(i+1)].sketches['__profile__'].AngularDimension(line1=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[27], line2=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[39], textPoint=(
        203.370040893555, -68.7961044311523), value=12.5)
    mdb.models['Model-%s'%(i+1)].sketches['__profile__'].AngularDimension(line1=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[36], line2=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[39], textPoint=(
        397.944091796875, -45.7543258666992), value=12.5)
    mdb.models['Model-%s'%(i+1)].sketches['__profile__'].ObliqueDimension(textPoint=(
        399.885040283203, -205.071746826172), value=10.0, vertex1=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].vertices[31], vertex2=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].vertices[32])
    mdb.models['Model-%s'%(i+1)].sketches['__profile__'].DistanceDimension(entity1=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[34], entity2=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[31], textPoint=(
        363.6474609375, -141.626968383789), value=59.0)
    mdb.models['Model-%s'%(i+1)].sketches['__profile__'].ObliqueDimension(textPoint=(
        748.884765625, -56.7823143005371), value=32.0, vertex1=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].vertices[40], vertex2=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].vertices[0])
    mdb.models['Model-%s'%(i+1)].sketches['__profile__'].VerticalDimension(textPoint=(
        718.648559570313, -62.8334770202637), value=52.5, vertex1=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].vertices[39], vertex2=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].vertices[0])
    mdb.models['Model-%s'%(i+1)].sketches['__profile__'].ObliqueDimension(textPoint=(
        773.009460449219, -98.5352630615234), value=25.0, vertex1=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].vertices[0], vertex2=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].vertices[1])
    mdb.models['Model-%s'%(i+1)].sketches['__profile__'].dragEntity(entity=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].vertices[43], points=((800.0, 
        0.0), (800.0, 0.0), (805.0, 0.0), (810.0, 0.0), (805.0, 0.0), (800.0, 0.0), 
        (795.0, 0.0), (795.0, 0.0)))
    mdb.models['Model-%s'%(i+1)].sketches['__profile__'].undo()
    mdb.models['Model-%s'%(i+1)].sketches['__profile__'].undo()
    mdb.models['Model-%s'%(i+1)].sketches['__profile__'].HorizontalConstraint(entity=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[5])
    mdb.models['Model-%s'%(i+1)].sketches['__profile__'].HorizontalConstraint(entity=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[7])
    mdb.models['Model-%s'%(i+1)].sketches['__profile__'].VerticalConstraint(entity=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[38])
    mdb.models['Model-%s'%(i+1)].sketches['__profile__'].VerticalConstraint(entity=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[6])
    mdb.models['Model-%s'%(i+1)].sketches['__profile__'].ObliqueDimension(textPoint=(
        772.044494628906, -99.7454986572266), value=25.0, vertex1=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].vertices[0], vertex2=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].vertices[1])
    mdb.models['Model-%s'%(i+1)].sketches['__profile__'].ObliqueDimension(textPoint=(
        807.7490234375, -55.572093963623), value=45.0, vertex1=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].vertices[1], vertex2=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].vertices[2])
    mdb.models['Model-%s'%(i+1)].sketches['__profile__'].ObliqueDimension(textPoint=(
        791.987548828125, -14.1217002868652), value=15.0, vertex1=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].vertices[2], vertex2=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].vertices[3])
    mdb.models['Model-%s'%(i+1)].sketches['__profile__'].VerticalConstraint(entity=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[34])
    mdb.models['Model-%s'%(i+1)].sketches['__profile__'].VerticalConstraint(entity=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[29])
    mdb.models['Model-%s'%(i+1)].sketches['__profile__'].VerticalConstraint(entity=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[31])
    mdb.models['Model-%s'%(i+1)].sketches['__profile__'].HorizontalConstraint(entity=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[30])
    mdb.models['Model-%s'%(i+1)].sketches['__profile__'].HorizontalConstraint(entity=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[32])
    mdb.models['Model-%s'%(i+1)].sketches['__profile__'].HorizontalConstraint(entity=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[26])
    mdb.models['Model-%s'%(i+1)].sketches['__profile__'].copyMirror(mirrorLine=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[39], objectList=(
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[5], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[6], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[7], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[26], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[27], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[28], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[29], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[30], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[31], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[32], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[33], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[34], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[35], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[36], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[37], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[38], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[40], 
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'].geometry[41]))
    mdb.models['Model-%s'%(i+1)].Part(dimensionality=AXISYMMETRIC, name='Part-1', type=
        DEFORMABLE_BODY)
    mdb.models['Model-%s'%(i+1)].parts['Part-1'].BaseShell(sketch=
        mdb.models['Model-%s'%(i+1)].sketches['__profile__'])
    del mdb.models['Model-%s'%(i+1)].sketches['__profile__']
    mdb.models['Model-%s'%(i+1)].ConstrainedSketch(name='__edit__', objectToCopy=
        mdb.models['Model-%s'%(i+1)].parts['Part-1'].features['Shell planar-1'].sketch)
    mdb.models['Model-%s'%(i+1)].parts['Part-1'].projectReferencesOntoSketch(filter=
        COPLANAR_EDGES, sketch=mdb.models['Model-%s'%(i+1)].sketches['__edit__'], 
        upToFeature=
        mdb.models['Model-%s'%(i+1)].parts['Part-1'].features['Shell planar-1'])
    mdb.models['Model-%s'%(i+1)].sketches['__edit__'].TangentConstraint(entity1=
        mdb.models['Model-%s'%(i+1)].sketches['__edit__'].geometry[35], entity2=
        mdb.models['Model-%s'%(i+1)].sketches['__edit__'].geometry[34])
    mdb.models['Model-%s'%(i+1)].sketches['__edit__'].TangentConstraint(entity1=
        mdb.models['Model-%s'%(i+1)].sketches['__edit__'].geometry[35], entity2=
        mdb.models['Model-%s'%(i+1)].sketches['__edit__'].geometry[36])
    mdb.models['Model-%s'%(i+1)].parts['Part-1'].features['Shell planar-1'].setValues(
        sketch=mdb.models['Model-%s'%(i+1)].sketches['__edit__'])
    del mdb.models['Model-%s'%(i+1)].sketches['__edit__']
    mdb.models['Model-%s'%(i+1)].parts['Part-1'].regenerate()
################################################################################################
    
    mdb.models['Model-%s'%(i+1)].Material(name='Material-1')
    mdb.models['Model-%s'%(i+1)].materials['Material-1'].Elastic(table=((E[i], v[i]), 
        ))
    mdb.models['Model-%s'%(i+1)].materials['Material-1'].Conductivity(table=((k[i], ), 
        ))
    mdb.models['Model-%s'%(i+1)].materials['Material-1'].Expansion(table=((a[i], ), ))
    mdb.models['Model-%s'%(i+1)].materials['Material-1'].Density(table=((7.85e-09, ), ))
    mdb.models['Model-%s'%(i+1)].HomogeneousSolidSection(material='Material-1', name=
        'Section-1', thickness=None)
    mdb.models['Model-%s'%(i+1)].parts['Part-1'].Set(faces=
        mdb.models['Model-%s'%(i+1)].parts['Part-1'].faces.getSequenceFromMask(('[#1 ]', 
        ), ), name='Set-1')
    mdb.models['Model-%s'%(i+1)].parts['Part-1'].SectionAssignment(offset=0.0, 
        offsetField='', offsetType=MIDDLE_SURFACE, region=
        mdb.models['Model-%s'%(i+1)].parts['Part-1'].sets['Set-1'], sectionName=
        'Section-1', thicknessAssignment=FROM_SECTION)
    mdb.models['Model-%s'%(i+1)].rootAssembly.DatumCsysByThreePoints(coordSysType=
        CYLINDRICAL, origin=(0.0, 0.0, 0.0), point1=(1.0, 0.0, 0.0), point2=(0.0, 
        0.0, -1.0))
    mdb.models['Model-%s'%(i+1)].rootAssembly.Instance(dependent=ON, name='Part-1-1', 
        part=mdb.models['Model-%s'%(i+1)].parts['Part-1'])
    mdb.models['Model-%s'%(i+1)].CoupledTempDisplacementStep(amplitude=RAMP, cetol=None, 
        creepIntegration=None, deltmx=None, name='Step-1', previous='Initial', 
        response=STEADY_STATE)
    mdb.models['Model-%s'%(i+1)].rootAssembly.Surface(name='Surf-1', side1Edges=
        mdb.models['Model-%s'%(i+1)].rootAssembly.instances['Part-1-1'].edges.getSequenceFromMask(
        ('[#3fffc0 ]', ), ))
    mdb.models['Model-%s'%(i+1)].FilmCondition(createStepName='Step-1', definition=
        EMBEDDED_COEFF, filmCoeff=Ht[i], filmCoeffAmplitude='', name='Int-1', 
        sinkAmplitude='', sinkDistributionType=UNIFORM, sinkFieldName='', 
        sinkTemperature=Tt[i], surface=
        mdb.models['Model-%s'%(i+1)].rootAssembly.surfaces['Surf-1'])
    mdb.models['Model-%s'%(i+1)].rootAssembly.Surface(name='Surf-2', side1Edges=
        mdb.models['Model-%s'%(i+1)].rootAssembly.instances['Part-1-1'].edges.getSequenceFromMask(
        ('[#ff00000f #f ]', ), ))
    mdb.models['Model-%s'%(i+1)].FilmCondition(createStepName='Step-1', definition=
        EMBEDDED_COEFF, filmCoeff=Hb[i], filmCoeffAmplitude='', name='Int-2', 
        sinkAmplitude='', sinkDistributionType=UNIFORM, sinkFieldName='', 
        sinkTemperature=Tb[i], surface=
        mdb.models['Model-%s'%(i+1)].rootAssembly.surfaces['Surf-2'])
    mdb.models['Model-%s'%(i+1)].rootAssembly.Set(edges=
        mdb.models['Model-%s'%(i+1)].rootAssembly.instances['Part-1-1'].edges.getSequenceFromMask(
        ('[#c00000 ]', ), ), name='Set-1')
    mdb.models['Model-%s'%(i+1)].DisplacementBC(amplitude=UNSET, createStepName='Step-1', 
        distributionType=UNIFORM, fieldName='', fixed=OFF, localCsys=None, name=
        'BC-1', region=mdb.models['Model-%s'%(i+1)].rootAssembly.sets['Set-1'], u1=0.0, 
        u2=0.0, ur3=UNSET)
    mdb.models['Model-%s'%(i+1)].rootAssembly.Set(name='Set-2', vertices=
        mdb.models['Model-%s'%(i+1)].rootAssembly.instances['Part-1-1'].vertices.getSequenceFromMask(
        ('[#800 ]', ), ))
    mdb.models['Model-%s'%(i+1)].ConcentratedForce(cf2=-F[i], createStepName='Step-1', 
        distributionType=UNIFORM, field='', localCsys=None, name='Load-1', region=
        mdb.models['Model-%s'%(i+1)].rootAssembly.sets['Set-2'])
    mdb.models['Model-%s'%(i+1)].rootAssembly.Set(faces=
        mdb.models['Model-%s'%(i+1)].rootAssembly.instances['Part-1-1'].faces.getSequenceFromMask(
        ('[#1 ]', ), ), name='Set-3')
    mdb.models['Model-%s'%(i+1)].RotationalBodyForce(centrifugal=ON, createStepName=
        'Step-1', magnitude=W[i], name='Load-2', point1=(0.0, 0.0, 0.0), point2=(
        0.0, 1.0, 0.0), region=mdb.models['Model-%s'%(i+1)].rootAssembly.sets['Set-3'], 
        rotaryAcceleration=OFF)
    mdb.models['Model-%s'%(i+1)].parts['Part-1'].seedPart(deviationFactor=0.1, 
        minSizeFactor=0.1, size=3.5)
    mdb.models['Model-%s'%(i+1)].parts['Part-1'].generateMesh()
    mdb.models['Model-%s'%(i+1)].parts['Part-1'].setElementType(elemTypes=(ElemType(
        elemCode=CAX4T, elemLibrary=STANDARD), ElemType(elemCode=CAX3T, 
        elemLibrary=STANDARD)), regions=(
        mdb.models['Model-%s'%(i+1)].parts['Part-1'].faces.getSequenceFromMask(('[#1 ]', 
        ), ), ))
    mdb.models['Model-%s'%(i+1)].rootAssembly.regenerate()
    mdb.Job(atTime=None, contactPrint=OFF, description='', echoPrint=OFF, 
        explicitPrecision=SINGLE, getMemoryFromAnalysis=True, historyPrint=OFF, 
        memory=90, memoryUnits=PERCENTAGE, model='Model-%s'%(i+1), modelPrint=OFF, 
        multiprocessingMode=DEFAULT, name='Thermomechanical-%s'%(i+1), nodalOutputPrecision=SINGLE, 
        numCpus=1, numGPUs=0, queue=None, resultsFormat=ODB, scratch='', type=
        ANALYSIS, userSubroutine='', waitHours=0, waitMinutes=0)
    mdb.jobs['Thermomechanical-'+str(i+1)].submit(consistencyChecking=OFF)
    mdb.jobs['Thermomechanical-'+str(i+1)].waitForCompletion()
    
    #####################################################################################
    #writing result
    import xlrd  
    #from openpyxl import Workbook
    #wbo = Workbook()
    #wso = wbo.create_sheet()
    # Import abaqus odb work related modules
    from odbAccess import *
    from abaqusConstants import *
    from odbMaterial import *
    from odbSection import *

    # open the odb file, paramatrically referred, as each imput row creates
    # a new odb file
    odb = openOdb(path='C:/Kannan/Suja/Thermomechanical-'+str(i+1)+'.odb')
    # Refer to the last frame for final results
    lastFrame = odb.steps['Step-1'].frames[-1]

    # Field output U captured in the variable displacement
    defor = lastFrame.fieldOutputs['U']
    fieldValues1 = defor.values
    Disp_list = [];
    for q in range(0,len(fieldValues1)):    
        Disp = fieldValues1[q].magnitude    
        Disp_list.append(Disp);
    Max_Disp = max(Disp_list);
    
    Mises = lastFrame.fieldOutputs['S']
    fieldValues2 = Mises.values
    Mises_stress_list = [];
    for q in range(0,len(fieldValues2)):    
        Mises_stress = fieldValues2[q].mises  
        Mises_stress_list.append(Mises_stress);
    Max_Mises_stress = max(Mises_stress_list);
    
    Mises = lastFrame.fieldOutputs['S']
    fieldValues3 = Mises.values
    Mises_stress_list_S11 = [];
    for q in range(0,len(fieldValues3)):    
        Mises_stress_S11 = fieldValues3[q].data[0]   # S11 = data[0]
        Mises_stress_list_S11.append(Mises_stress_S11);
    Max_Stress_S11 = max(Mises_stress_list_S11);
    
    Mises = lastFrame.fieldOutputs['S']
    fieldValues4 = Mises.values
    Mises_stress_list_S22 = [];
    for q in range(0,len(fieldValues4)):    
        Mises_stress_S22 = fieldValues4[q].data[1]    # S22 = data[1]
        Mises_stress_list_S22.append(Mises_stress_S22);
    Max_Stress_S22 = max(Mises_stress_list_S22);
    
    Mises = lastFrame.fieldOutputs['HFL']
    fieldValues5 = Mises.values
    Mises_disp_list_HFL = [];
    for q in range(0,len(fieldValues5)):    
        Mises_HFL = fieldValues5[q].magnitude    
        Mises_disp_list_HFL.append(Mises_HFL);
    Max_HFL = max(Mises_disp_list_HFL);
    
    #print 'Stress magnitude =', Max_Mises_stress
    wso.write(i,0,Max_Mises_stress)
    wso.write(i,1,Max_Disp)
    wso.write(i,2,Max_Stress_S11)
    wso.write(i,3,Max_Stress_S22)
    wso.write(i,4,Max_HFL)
workbook.close()
    #####################################################################################
