# 见习实训 — Cesium 三维地球可视化项目

黄冈师范学院 计算机学院 数据科学与大数据技术 见习实训课程项目。

## 项目概述

基于 CesiumJS 构建的三维地球 WebGIS 应用，包含以下实验模块：

| 实验 | 文件 | 功能 |
|------|------|------|
| 实验1-1 | `exp1_helloworld.html` | 构建 Cesium 三维地球场景，禁用默认 UI 控件，ArcGIS 卫星影像底图 |
| 实验1-2 | `exp2_box_entity.html` | 三维场景中添加立方体要素、椭圆实体，按钮交互控制 |
| 实验2-1 | `exp3_pointcloud.html` | 加载 3D Tiles 点云数据，设置红色样式，添加圆形区域标记 |
| 实验2-2 | `exp4_model_rotation.html` | 加载 3D Tiles 三维模型，绕 X/Y/Z 轴旋转交互 |

## 技术栈

- **三维引擎**: CesiumJS 1.40
- **底图服务**: ArcGIS Satellite Imagery / SingleTileImageryProvider
- **数据格式**: 3D Tiles (.b3dm, .i3dm, .pnts)
- **本地服务器**: Node.js 零依赖静态文件服务器 (`server.js`)
- **前端**: 原生 HTML5 + JavaScript + CSS3

## 项目结构

```
Internship-Project-Title/
├── server.js                    # 零依赖本地 HTTP 服务器
├── exp1_helloworld.html         # 实验1-1：3D 地球场景
├── exp2_box_entity.html         # 实验1-2：立方体要素
├── exp3_pointcloud.html         # 实验2-1：点云数据
├── exp4_model_rotation.html     # 实验2-2：模型旋转
├── 见习题目/
│   ├── Cesium-1.40/             # CesiumJS 库（需自行下载放置）
│   │   ├── Build/Cesium/        # 编译后的 JS/CSS/Workers
│   │   └── img/worldimage.jpg   # 世界影像底图
│   ├── pwt1/pts/                # 点云 3D Tiles 数据
│   ├── Tileset/                 # 模型 3D Tiles 数据
│   └── 3dtilesDemo/             # 大型建筑模型数据
└── .gitignore
```

## 快速开始

### 1. 环境要求

- Node.js v14+
- 现代浏览器（Chrome / Edge / Firefox），需支持 WebGL

### 2. 启动服务器

```bash
cd Internship-Project-Title
node server.js
```

### 3. 访问页面

```
http://localhost:8080/exp1_helloworld.html    # 实验1-1
http://localhost:8080/exp2_box_entity.html    # 实验1-2
http://localhost:8080/exp3_pointcloud.html    # 实验2-1
http://localhost:8080/exp4_model_rotation.html # 实验2-2
```

> **注意**：CesiumJS 库文件较大（约 50MB），未包含在 Git 仓库中。请从 [CesiumJS 官网](https://cesium.com/platform/cesiumjs/) 下载后放置到 `见习题目/Cesium-1.40/` 目录。

## 实验截图

*（打开页面后可通过浏览器截图工具捕获）*

### 实验1-1：三维地球场景
![exp1_helloworld](screenshots/exp1_helloworld.png)

### 实验1-2：立方体要素
![exp2_box_entity](screenshots/exp2_box_entity.png)

### 实验2-1：点云数据
![exp3_pointcloud](screenshots/exp3_pointcloud.png)

### 实验2-2：模型旋转
![exp4_model_rotation](screenshots/exp4_model_rotation.png)

## 常见问题

| 问题 | 解决方案 |
|------|----------|
| 页面空白 / Cesium 未定义 | 确保 Cesium-1.40 放在 `见习题目/` 下，路径完整 |
| 点云/模型加载失败 | 检查 `见习题目/pwt1/pts/` 和 `见习题目/Tileset/` 目录存在 |
| 按钮点击无反应 | 打开浏览器控制台查看错误信息 |
| 端口被占用 | `netstat -ano \| findstr ":8080"` 然后 `taskkill /F /PID <PID>` |

## 许可

本项目仅用于教育学习目的。
