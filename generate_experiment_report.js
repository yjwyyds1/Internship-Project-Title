const fs = require('fs');
const { Document, Packer, Paragraph, TextRun, HeadingLevel,
    AlignmentType, PageBreak, BorderStyle, WidthType, LevelFormat,
    PageNumber, Header, Footer, PageOrientation, TabStopType, TabStopPosition,
    TableOfContents } = require('docx');

function createExperimentReport() {
    const doc = new Document({
        styles: {
            default: { document: { run: { font: "SimSun", size: 24 } } },
            paragraphStyles: [
                { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
                    run: { size: 32, bold: true, font: "SimHei" },
                    paragraph: { spacing: { before: 240, after: 240 }, outlineLevel: 0, alignment: AlignmentType.CENTER } },
                { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
                    run: { size: 28, bold: true, font: "SimHei" },
                    paragraph: { spacing: { before: 180, after: 180 }, outlineLevel: 1 } },
                { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
                    run: { size: 26, bold: true, font: "SimHei" },
                    paragraph: { spacing: { before: 120, after: 120 }, outlineLevel: 2 } },
            ]
        },
        numbering: {
            config: [
                { reference: "bullets",
                    levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u25CF", alignment: AlignmentType.LEFT,
                        style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
            ]
        },
        sections: [
            // Cover page
            {
                properties: {
                    page: {
                        size: { width: 11906, height: 16838 },
                        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
                    }
                },
                children: [
                    new Paragraph({ spacing: { after: 200 }, children: [] }),
                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { after: 100 },
                        children: [new TextRun({ text: "黄冈师范学院", size: 56, bold: true, font: "SimHei" })]
                    }),
                    new Paragraph({ spacing: { after: 400 }, children: [] }),
                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { after: 200 },
                        children: [new TextRun({ text: "见习实训", size: 48, bold: true, font: "SimHei" })]
                    }),
                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { after: 200 },
                        children: [new TextRun({ text: "课程实验报告", size: 40, bold: true, font: "SimHei" })]
                    }),
                    new Paragraph({ spacing: { after: 400 }, children: [] }),
                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { after: 200 },
                        children: [new TextRun({ text: "院\t系：计算机学院", size: 28, font: "SimSun" })]
                    }),
                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { after: 200 },
                        children: [new TextRun({ text: "专\t业：数据科学与大数据技术", size: 28, font: "SimSun" })]
                    }),
                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { after: 200 },
                        children: [new TextRun({ text: "课\t程：见习实训", size: 28, font: "SimSun" })]
                    }),
                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { after: 200 },
                        children: [new TextRun({ text: "学生姓名：___________", size: 28, font: "SimSun" })]
                    }),
                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { after: 200 },
                        children: [new TextRun({ text: "班\t级：大数据2301", size: 28, font: "SimSun" })]
                    }),
                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { after: 200 },
                        children: [new TextRun({ text: "学\t号：2023310140228", size: 28, font: "SimSun" })]
                    }),
                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { after: 200 },
                        children: [new TextRun({ text: "学\t期：2025-2026第二学期", size: 28, font: "SimSun" })]
                    }),
                ]
            },
            // Experiment 1
            createExperimentSection(
                "实验1：构建Cesium三维地球场景",
                "1. 安装、配置并使用Cesium进行三维地球场景的可视化\n" +
                "2. 了解Cesium.Viewer的基本配置，掌握如何控制UI控件\n" +
                "3. 理解Cesium在Web端的地理空间可视化原理\n" +
                "4. 熟悉Node.js本地服务器环境的搭建",
                "步骤1：下载CesiumJS库并解压到项目文件夹\n\n" +
                "步骤2：创建HTML文件，引入Cesium核心库（Cesium.js）和样式文件（widgets.css）\n\n" +
                "步骤3：在页面中创建div容器用于Cesium渲染3D场景\n\n" +
                "步骤4：使用Cesium.Viewer初始化3D地球场景，设置ArcGIS影像底图\n\n" +
                "步骤5：禁用不需要的默认控件（如geocoder、homeButton、animation等），简化界面\n\n" +
                "步骤6：设置相机视角飞到中国区域\n\n" +
                "核心代码：\n" +
                "var viewer = new Cesium.Viewer('cesiumContainer', {\n" +
                "    imageryProvider: new Cesium.ArcGisMapServerImageryProvider({...}),\n" +
                "    geocoder: false, homeButton: false, sceneModePicker: false,\n" +
                "    baseLayerPicker: false, animation: false, timeline: false,\n" +
                "    fullscreenButton: false, vrButton: false\n" +
                "});\n\n" +
                "运行方式：使用Node.js启动本地服务器，访问http://localhost:8080/exp1_helloworld.html\n\n" +
                "运行结果：浏览器成功加载3D地球场景，显示卫星影像底图，界面简洁，可通过鼠标进行旋转、缩放和倾斜操作。",
                "通过本次实验，我掌握了以下内容：\n" +
                "1. 如何在HTML中引入并使用Cesium库，理解其模块化加载方式\n" +
                "2. 了解了Cesium.Viewer的基本配置选项，包括影像底图设置和UI控件控制\n" +
                "3. 熟悉了Node.js本地服务器的搭建方法，解决了跨域和静态资源加载问题\n" +
                "4. 初步了解了WebGIS在浏览器端的展示方式和渲染原理\n\n" +
                "遇到的问题与解决：\n" +
                "- Cesium.js和widgets.css路径引用错误导致页面空白，通过检查Network请求找到正确相对路径\n" +
                "- Bing Maps的AccessToken失效导致底图无法加载，改用ArcGIS影像服务解决\n" +
                "- 本地文件跨域问题需要使用HTTP服务器（如node server.js）而非直接打开HTML文件\n\n" +
                "改进建议：可以尝试加载更多类型的底图数据（如OpenStreetMap、天地图等），以及添加交互功能"
            ),
            // Experiment 2
            createExperimentSection(
                "实验2：Cesium三维场景中添加立方体要素",
                "1. 掌握如何通过Cesium的Entity API向三维场景中添加几何实体\n" +
                "2. 理解实体的位置、尺寸、材质等属性的设置方法\n" +
                "3. 学会使用zoomTo方法将相机视角聚焦到特定实体\n" +
                "4. 掌握实体的动态添加与移除操作",
                "步骤1：在已有Cesium Viewer的基础上，使用viewer.entities.add()添加立方体实体\n\n" +
                "步骤2：设置实体的位置（使用Cartesian3.fromDegrees指定经纬度和高度）\n\n" +
                "步骤3：配置立方体属性，包括尺寸（dimensions）、材质（material）、轮廓（outline）等\n\n" +
                "步骤4：使用viewer.zoomTo()将相机聚焦到添加的实体上\n\n" +
                "步骤5：添加按钮控件实现动态添加、飞入和移除立方体的交互功能\n\n" +
                "核心代码：\n" +
                "var redBox = viewer.entities.add({\n" +
                "    name: 'Red box with black outline',\n" +
                "    position: Cesium.Cartesian3.fromDegrees(-107.0, 40.0, 300000.0),\n" +
                "    box: {\n" +
                "        dimensions: new Cesium.Cartesian3(400000.0, 300000.0, 500000.0),\n" +
                "        material: Cesium.Color.RED.withAlpha(0.5),\n" +
                "        outline: true,\n" +
                "        outlineColor: Cesium.Color.BLACK\n" +
                "    }\n" +
                "});\n\n" +
                "运行结果：页面加载后显示3D地球场景，点击'添加红色立方体'按钮后，在指定经纬度位置（美国西部）\n" +
                "出现一个半透明的红色立方体，带有黑色轮廓线，相机自动飞入到立方体所在位置。\n" +
                "可通过'移除立方体'按钮清除场景中的立方体实体。",
                "通过本次实验，我掌握了以下内容：\n" +
                "1. Cesium Entity API的基本使用方法，能够通过代码添加3D几何实体\n" +
                "2. 理解了Cartesian3坐标系的概念，学会了经纬度到笛卡尔坐标的转换\n" +
                "3. 掌握了立方体实体的尺寸、颜色、透明度、轮廓等属性的配置\n" +
                "4. 学习了实体操作函数（添加、移除、聚焦），为后续交互功能开发打下了基础\n\n" +
                "遇到的问题与解决：\n" +
                "- 立方体位置设置在地球内部导致看不到，需要将高度设为300000m以上\n" +
                "- 材质透明度设置不当导致立方体完全不可见，需确保withAlpha设置合理\n\n" +
                "改进建议：可以尝试添加其他几何类型（如球体、圆柱体、多边形等），并实现点击实体触发事件的功能"
            ),
            // Experiment 3
            createExperimentSection(
                "实验3：加载点云数据及添加圆形区域",
                "1. 掌握Cesium中加载3DTiles点云数据的基本方法\n" +
                "2. 学会为点云数据设置样式（颜色、点大小等）\n" +
                "3. 理解SingleTileImageryProvider的使用场景\n" +
                "4. 学会使用viewer.entities.add添加圆形（椭圆）实体\n" +
                "5. 掌握相机飞行（flyTo）的使用方法",
                "步骤1：创建HTML文件，初始化Viewer，使用SingleTileImageryProvider设置自定义底图\n\n" +
                "步骤2：使用viewer.scene.primitives.add加载pwt1/pts/tileset.json点云数据\n\n" +
                "步骤3：在tileset的readyPromise中为点云设置样式：统一颜色为红色（#ff0000），点大小设为5.0\n\n" +
                "步骤4：使用flyTo方法将相机飞行到指定位置（广东佛山区域的电力线点云数据）\n\n" +
                "步骤5：使用viewer.entities.add添加两个蓝色圆形标记，指定半轴长度\n\n" +
                "步骤6：添加交互按钮实现飞入点云区域、切换点云显示、添加圆形标记等功能\n\n" +
                "核心代码：\n" +
                "var pointsTileset = viewer.scene.primitives.add(\n" +
                "    new Cesium.Cesium3DTileset({ url: 'pwt1/pts/tileset.json' })\n" +
                ");\n" +
                "pointsTileset.style = new Cesium.Cesium3DTileStyle({\n" +
                "    color: \"color('#ff0000')\", pointSize: 5.0\n" +
                "});\n" +
                "viewer.entities.add({\n" +
                "    position: Cesium.Cartesian3.fromDegrees(112.82536, 23.071506, 100),\n" +
                "    ellipse: {\n" +
                "        semiMinorAxis: 300.0, semiMajorAxis: 300.0,\n" +
                "        material: Cesium.Color.BLUE.withAlpha(1.0)\n" +
                "    }\n" +
                "});\n\n" +
                "运行结果：页面加载后使用自定义世界影像底图，自动加载红色点云数据（广东佛山220KV高压线\n" +
                "路点云），相机飞行到点云区域（经度112.82536，纬度23.071506），可通过按钮切换点云显隐\n" +
                "和添加蓝色圆形标记。",
                "通过本次实验，我掌握了以下内容：\n" +
                "1. 学会了使用Cesium3DTileset加载3DTiles点云数据（.pnts格式）的方法\n" +
                "2. 理解了Cesium3DTileStyle的作用，能够自定义点云的颜色和点大小\n" +
                "3. 掌握了SingleTileImageryProvider的使用方法，可以使用单张图片作为底图\n" +
                "4. 学会了使用实体API绘制圆形（椭圆）图形标记\n" +
                "5. 掌握了flyTo方法的各种参数配置（destination、orientation等）\n\n" +
                "遇到的问题与解决：\n" +
                "- 点云数据显示为黑色或不可见，需要在readyPromise中设置样式后才能正常渲染\n" +
                "- 点云数据文件路径错误，需确保tileset.json及相关文件在正确目录下\n" +
                "- pointSize参数在部分点云数据上可能无效，需要使用支持该属性的数据源\n" +
                "- 底图图片引用失败，改用ArcGIS在线影像服务作为备用方案\n\n" +
                "改进建议：可以尝试使用更多样式属性（如按高程着色、按类别分类显示等），\n" +
                "以及添加GeoJSON电力线数据进行叠加显示"
            ),
            // Experiment 4
            createExperimentSection(
                "实验4：加载3DTiles模型并实现旋转交互",
                "1. 掌握Cesium中加载3DTiles模型数据的方法\n" +
                "2. 理解tileset的modelMatrix和矩阵变换原理\n" +
                "3. 学习通过Matrix3和Matrix4实现模型的旋转操作\n" +
                "4. 掌握按钮交互控制模型旋转的方法\n" +
                "5. 理解3DTiles的数据结构（tileset.json、b3dm瓦片文件等）",
                "步骤1：初始化Viewer，使用ArcGIS卫星影像作为底图\n\n" +
                "步骤2：使用viewer.scene.primitives.add加载Tileset/tileset.json三维模型数据\n\n" +
                "步骤3：在模型readyPromise中获取原始模型矩阵（modelMatrix）并保存\n\n" +
                "步骤4：相机自动飞入模型所在位置\n\n" +
                "步骤5：添加三个旋转按钮（绕X轴、Y轴、Z轴旋转），每次点击旋转1度（累积式旋转）\n\n" +
                "步骤6：实现旋转函数，使用Matrix3.fromRotationX/Y/Z计算旋转矩阵，\n" +
                "      再使用Matrix4.multiplyByMatrix3将旋转矩阵应用到模型矩阵\n\n" +
                "步骤7：添加重置按钮，将模型恢复到初始姿态\n\n" +
                "步骤8：在界面上显示当前各轴的累积旋转角度\n\n" +
                "核心代码：\n" +
                "var tileset = viewer.scene.primitives.add(\n" +
                "    new Cesium.Cesium3DTileset({ url: 'Tileset/tileset.json' })\n" +
                ");\n" +
                "var originalMatrix = Cesium.Matrix4.clone(tileset.modelMatrix);\n" +
                "function rotateX() {\n" +
                "    anglex += 1;\n" +
                "    var rotMat = Cesium.Matrix3.fromRotationX(Cesium.Math.toRadians(anglex));\n" +
                "    tileset.modelMatrix = Cesium.Matrix4.multiplyByMatrix3(\n" +
                "        originalMatrix, rotMat, new Cesium.Matrix4()\n" +
                "    );\n" +
                "}\n\n" +
                "运行结果：页面加载后自动加载3DTiles模型（由1个父级b3dm和4个子级b3dm组成），\n" +
                "相机自动飞入模型位置。通过'绕X轴旋转'、'绕Y轴旋转'、'绕Z轴旋转'三个按钮\n" +
                "可以实时控制模型在各轴上的旋转，界面底部实时显示当前旋转角度。点击'重置旋转'\n" +
                "可以将模型恢复到初始姿态。",
                "通过本次实验，我掌握了以下内容：\n" +
                "1. 学会了使用Cesium3DTileset加载3DTiles三维模型（.b3dm格式）数据\n" +
                "2. 深入理解了Cesium中矩阵变换（Matrix3和Matrix4）的原理和应用\n" +
                "3. 掌握了通过modelMatrix实现模型旋转、平移等空间变换的方法\n" +
                "4. 理解了3DTiles的数据结构，包括tileset.json的配置、b3dm瓦片文件的组织方式\n" +
                "5. 学会了创建交互按钮与Cesium场景进行实时交互\n\n" +
                "遇到的问题与解决：\n" +
                "- 旋转后模型位置偏移，原因是每次旋转都在原矩阵上累积，应始终基于originalMatrix进行旋转\n" +
                "- 旋转角度单位错误（将度直接传给fromRotationX），需要使用Math.toRadians转换为弧度\n" +
                "- 模型加载失败时需要检查tileset.json中的url路径是否正确指向b3dm文件\n" +
                "- 多个旋转同时使用时需要分别追踪各轴的角度（anglex、angley、anglez）\n\n" +
                "改进建议：可以增加平移和缩放功能，实现完整的模型浏览交互；\n" +
                "也可以添加鼠标拖拽旋转和滚轮缩放等更直观的交互方式"
            ),
        ]
    });

    const outputPath = 'D:/Jxpython/Internship-Project-Title/课程实验报告.docx';
    Packer.toBuffer(doc).then(buffer => {
        fs.writeFileSync(outputPath, buffer);
        console.log('Experiment report created: ' + outputPath);
    });
}

function createExperimentSection(title, purpose, procedure, results) {
    return {
        properties: {
            page: {
                size: { width: 11906, height: 16838 },
                margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
            }
        },
        headers: {
            default: new Header({
                children: [new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [new TextRun({ text: "黄冈师范学院 见习实训 课程实验报告", size: 18, font: "SimSun", color: "888888" })]
                })]
            })
        },
        footers: {
            default: new Footer({
                children: [new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [new TextRun({ text: "第 ", size: 18 }), new TextRun({ children: [PageNumber.CURRENT], size: 18 }), new TextRun({ text: " 页", size: 18 })]
                })]
            })
        },
        children: [
            new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(title)] }),
            new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 200 },
                border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "2E75B6", space: 1 } },
                children: []
            }),
            new Paragraph({
                heading: HeadingLevel.HEADING_2,
                children: [new TextRun("一、实验课题")]
            }),
            new Paragraph({
                spacing: { after: 200 },
                children: [new TextRun({ text: title.replace(/^实验\d[：:]/, "").trim(), size: 26, font: "SimHei" })]
            }),
            new Paragraph({
                heading: HeadingLevel.HEADING_2,
                children: [new TextRun("二、实验目的及要求")]
            }),
            ...purpose.split('\n').filter(l => l.trim()).map(line =>
                new Paragraph({
                    spacing: { after: 60 },
                    children: [new TextRun({ text: line, size: 24, font: "SimSun" })]
                })
            ),
            new Paragraph({
                heading: HeadingLevel.HEADING_2,
                children: [new TextRun("三、实验步骤及结果")]
            }),
            ...procedure.split('\n').filter(l => l.trim()).map(line =>
                new Paragraph({
                    spacing: { after: 60 },
                    children: [new TextRun({ text: line, size: 24, font: "SimSun" })]
                })
            ),
            new Paragraph({
                heading: HeadingLevel.HEADING_2,
                children: [new TextRun("四、实验结果及体会")]
            }),
            ...results.split('\n').filter(l => l.trim()).map(line =>
                new Paragraph({
                    spacing: { after: 60 },
                    children: [new TextRun({ text: line, size: 24, font: "SimSun" })]
                })
            ),
            new Paragraph({ children: [new PageBreak()] }),
        ]
    };
}

createExperimentReport();
