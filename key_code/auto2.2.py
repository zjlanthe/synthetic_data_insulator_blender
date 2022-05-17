# encoding:utf-8
import bpy
import os
import math
import random
import time
import xlwt  # pip install xlwt  需安装

C = bpy.context
scn = C.scene

Mobiledistance = 0.16  # 绝缘子间距
Num = 2  # 增加的绝缘子数量
Zoom_ratio = 1  # 缩放比例
mobile_x = 0    # 模型位置
mobile_y = 0
mobile_z = 0.3
rotating_x = -1  # 模型角度
rotating_y = 1
rotating_z = 0

# Get the environment node tree of the current scene
node_tree = bpy.context.scene.world.node_tree
tree_nodes = node_tree.nodes

# Fi0 = os.getcwd() + "\\" + "sysdata" + "\\" + "JPEGImages" + "\\"
# Fi1 = os.getcwd() + "\\" + "sysdata" + "\\" + "Annotations" + "\\"
# Fi2 = os.getcwd() + "\\" + "sysdata" + "\\" + "Image_no_Annotations" + "\\"
# Fi3 = os.getcwd() + "\\" + "sysdata" + "\\" + "timedata" + "\\"
# savepath = Fi3 + 'rendertime.xls'

F_folder = bpy.path.abspath("//sysdata")
Fi0 = F_folder + "\\" + "JPEGImages" + "\\"
Fi1 = F_folder + "\\" + "Annotations" + "\\"
Fi2 = F_folder + "\\" + "Image_no_Annotations" + "\\"
Fi3 = F_folder + "\\" + "timedata" + "\\"
savepath = Fi3 + 'excel.xls'


# 创建文件夹
def mkdir_dataset(File_Path):
    try:
        print(File_Path)
        if not os.path.exists(File_Path):
            os.makedirs(File_Path)
            print("succeed" + File_Path)
        else:
            print("already exist")
    except BaseException as msg:
        print("Failed to create folder")


# 缩放 组装
def zoom():
    bpy.ops.object.select_camera()  # 选择相机
    bpy.ops.object.select_all(action='INVERT')  # 反选

    bpy.ops.object.origin_set(type='ORIGIN_CURSOR', center='MEDIAN')
    bpy.ops.transform.resize(value=(Zoom_ratio, Zoom_ratio, Zoom_ratio))  # 缩放

    # 游标位置
    bpy.context.scene.cursor.location[0] = 0
    bpy.context.scene.cursor.location[1] = 0
    bpy.context.scene.cursor.location[2] = 0

    # copy
    for i in range(Num):

        bpy.ops.object.duplicate_move(TRANSFORM_OT_translate={
                                      "value": (0, 0, Mobiledistance)})
        bpy.context.selected_objects[0].name = "Insulation"+str(i)
        bpy.context.selected_objects[1].name = "fitting"+str(i)

    # combination
    bpy.ops.object.select_all(action='SELECT')  # SELECT 全选  DESELECT 取消选择
    # 重命名
    bpy.context.view_layer.objects.active = bpy.data.objects["Insulation"]
    bpy.ops.object.join()


# mobile
def move():
    # 取消选择
    bpy.ops.object.select_all(action='DESELECT')
    # 移动位置
    bpy.data.objects['Insulation'].location.x = random.choice([-1.9, 1.8])
    bpy.data.objects['Insulation'].location.y = random.randrange(2, 6, 1)
    bpy.data.objects['Insulation'].location.z = random.choice([2.9, 0.4])
    # 旋转
    ran1 = random.randrange(-20, 20, 1)
    ran2 = random.randrange(60, 90, 1)
    ran3 = random.randrange(-90, -60, 1)
    ran = random.choice([ran1, ran2, ran3])  # 选择ran1、2、3其中一个
    bpy.data.objects['Insulation'].rotation_euler = (
        math.radians(ran), math.radians(random.randrange(150, 210, 1)), 0)

    # 选择相机
    bpy.ops.object.select_all(action='DESELECT')
    # 设置相机位置
    bpy.data.objects['摄像机'].location.x = 0.0
    bpy.data.objects['摄像机'].location.y = 0
    bpy.data.objects['摄像机'].location.z = 4
    # 设置相机角度
    x = bpy.data.objects['Insulation'].location.x
    y = bpy.data.objects['Insulation'].location.y
    z = bpy.data.objects['Insulation'].location.z
    xx = math.atan(math.sqrt(x*x+y*y)/(4 - z)) + \
        random.uniform(-math.pi/12, math.pi/12)
    zz = math.acos(((4 - z)*(4 - z)+y*math.sqrt(x*x+y*y)) / (x*x +
                   (4 - z)*(4 - z)+y*y)) + random.uniform(-math.pi/12, math.pi/12)
    if x > 0:
        bpy.context.scene.camera.rotation_euler = (xx, 0, -zz)
    else:
        bpy.context.scene.camera.rotation_euler = (xx, 0, zz)


def hdri(hdris):
    # Clear all nodes 清除所有节点
    tree_nodes.clear()

    # Add Background node 添加背景节点
    node_background = tree_nodes.new(type='ShaderNodeBackground')

    # Add Environment Texture node 添加环境节点
    node_environment = tree_nodes.new(type='ShaderNodeTexEnvironment')

    # Load and assign the image to the node property 加载hdri
    node_environment.image = bpy.data.images.load(hdris)
    node_environment.location = -300, 0  # 节点位置

    # 添加映射节点  （改变hdri方向）
    node_mapping = tree_nodes.new(type='ShaderNodeMapping')
    node_mapping.location = -500, 0

    # 添加纹理坐标节点  （改变hdri方向）
    node_coord = tree_nodes.new(type='ShaderNodeTexCoord')
    node_coord.location = -700, 0

    # Add Output node 添加输出节点
    node_output = tree_nodes.new(type='ShaderNodeOutputWorld')
    node_output.location = 200, 0

    # Link all nodes 节点连接
    links = node_tree.links
    link1 = links.new(
        node_environment.outputs["Color"], node_background.inputs["Color"])
    link2 = links.new(
        node_background.outputs["Background"], node_output.inputs["Surface"])
    link3 = links.new(
        node_mapping.outputs["Vector"], node_environment.inputs["Vector"])
    link4 = links.new(
        node_coord.outputs["Generated"], node_mapping.inputs["Vector"])

    return node_environment, node_background, link1


def hdri_adjust():
    #  改变hdri方向
    bpy.data.worlds["World"].node_tree.nodes["映射"].inputs[2].default_value[0] \
        = math.radians(random.randrange(-30, 30, 1))
    bpy.data.worlds["World"].node_tree.nodes["映射"].inputs[2].default_value[1] \
        = math.radians(random.randrange(-10, 10, 1))
    bpy.data.worlds["World"].node_tree.nodes["映射"].inputs[2].default_value[2] \
        = math.radians(random.randrange(0, 360, 1))


def light():
    # 调节亮度 默认1
    zoom_light = 0.1 * random.randrange(3, 10, 1)
    bpy.data.worlds["World"].node_tree.nodes["背景"].inputs[1].default_value \
        = zoom_light
    return zoom_light


def clear(node_environment, node_background, link1):
    # 清除hdri节点link1（背景-环境）
    bpy.context.scene.world.node_tree.links.remove(link1)


def save(file):
    # 渲染及保存图像
    scene = bpy.context.scene
    scene.render.image_settings.file_format = 'PNG'
    scene.render.filepath = file
    bpy.ops.render.render(write_still=1)


#  将数据写入新文件
def data_write(file_path, datas):
    f = xlwt.Workbook()
    sheet1 = f.add_sheet(u'sheet1', cell_overwrite_ok=True)  # 创建sheet

    # 将数据写入第 i 行，第 j 列
    i = 0
    for data in datas:
        for j in range(len(data)):
            sheet1.write(i, j, data[j])
        i = i + 1
    f.save(file_path)  # 保存文件


def cv_label():
    a = 1
    # opencv label标注另行运行（label_opencv.py）


def main():
    num_image = 0
    # 1280x720
    bpy.data.scenes["Scene"].render.resolution_x = 720
    bpy.data.scenes["Scene"].render.resolution_y = 1280
    # Get the HDRI
    hdri_folder = bpy.path.abspath("//hdri_based")
    hdri_file = [os.path.join(hdri_folder, f) for f in os.listdir(hdri_folder) if f.endswith(".exr")]
    # 创建文件处
    mkdir_dataset(Fi0)
    mkdir_dataset(Fi1)
    mkdir_dataset(Fi2)
    mkdir_dataset(Fi3)
    # 创建excel
    book = xlwt.Workbook(encoding='utf-8', style_compression=0)
    sheet = book.add_sheet('01', cell_overwrite_ok=True)
    # 正式运行
    for i in range(1, 2):  # 组装个数
        zoom()  # 组装
        for hdris in hdri_file:
            for k in range(1, 18):
                num_image += 1  # 循环次数计数
                total_time = 0  # 循环运行时间
                print(num_image)
                start = time.time()
                node_environment, node_background, link1 = hdri(hdris)  # 插入环境
                hdri_adjust()  # 调节hdri角度
                move()  # 移动
                zoom_light = light()  # 调节灯光
                end = time.time()
                sheet.write(num_image, 1, (end-start))
                total_time += (end-start)
                sheet.write(num_image, 4, zoom_light)

                # 保存有背景图像
                bpy.context.scene.render.engine = 'CYCLES'
                bpy.context.scene.cycles.device = 'GPU'
                start = time.time()
                save(Fi0 + str(num_image) + ".png")  # 保存
                end = time.time()
                sheet.write(num_image, 2, (end-start))
                total_time += (end-start)

                # 保存无背景图像
                bpy.context.scene.render.engine = 'BLENDER_EEVEE'
                bpy.data.worlds["World"].node_tree.nodes["背景"].inputs[1].default_value = 1    # 统一亮度为1
                start = time.time()
                clear(node_environment, node_background, link1)  # 清除link1节点（背景-环境）
                save(Fi2 + str(num_image) + ".png")  # 保存无背景图像
                end = time.time()
                sheet.write(num_image, 3, (end-start))
                total_time += (end-start)
                bpy.context.scene.world.node_tree.nodes.clear()

                print("循环运行时间:%.2f秒" % total_time)
                book.save(savepath)  # 保存excel
                # print(savepath)

                # 清除缓存（图像和hdri）
                # Clean up unused images
                for img in bpy.data.images:
                    # if not img.users:
                    bpy.data.images.remove(img)
                # Reload all File images
                for img in bpy.data.images:
                    if img.source == 'FILE':
                        img.reload()
    # book.save(savepath)
    # print(savepath)
    cv_label()


if __name__ == '__main__':
    main()
