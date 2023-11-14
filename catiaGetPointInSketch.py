import win32com.client
from win32com.client import Dispatch, VARIANT
import pythoncom

# Ensure that Python arrays are converted to a format that COM understands
pythoncom.CoInitialize()

# Create a VARIANT as a placeholder for the coordinates
# point_coordinates = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [0.0, 0.0])

catia = Dispatch("CATIA.Application").ActiveDocument
# print(dir(catia))

part2 = catia.Part
body=part2.HybridBodies[0]
print(body.name)
sketch= body.HybridSketches[2]
print(sketch.name)
part2.InWorkObject = sketch
factory2d = sketch.OpenEdition()
# point1=sketch.GeometricElements.Item("点.2")
constraint1 = sketch.Constraints.Item('y1')
y2 = sketch.Constraints.Item('y2')
print(y2)
a = constraint1.Dimension.Value
print(a)
print(type(a))
constraint1.Dimension.Value=360

coordinateDic = {}
for constraint in sketch.Constraints:
    print(constraint.name)






# print(dir(point1))
# print(point1_coords)

# hybridBodies1 = part2.HybridBodies
# hybridBody1 = hybridBodies1.Item("几何图形集.1")
# sketches1 = hybridBody1.HybridSketches
# sketch1 = sketches1.Item("y0")
#
# part2.InWorkObject = sketch1
# factory2D1 = sketch1.OpenEdition()
# geometricElements1 = sketch1.GeometricElements
# point2D1 = geometricElements1.Item("p3")
#
#
# # 使用 VT_R8 类型，这表示数组中的元素是双精度浮点数
# coords = [0.0, 0.0]# This specifies that you are passing a safe array of doubles
# # coords = win32com.client.VARIANT(pythoncom.VT_VARIANT | pythoncom.VT_ARRAY, [0.0, 0.0])
#
# coords = point2D1.GetCoordinates(coords)
# print(coords)
# x_value = coords[0]
# y_value = coords[1]
# print(x_value, y_value)
#
# point2D1.SetData(1100.0, 1000.0)
part2.UpdateObject(y2)


sketch.CloseEdition()
part2.Update()
del catia

