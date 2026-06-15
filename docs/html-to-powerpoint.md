<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, user-scalable=yes">
    <title>商场月度店长会议PPT生成器</title>
    <!-- 引入 pptxgen.js 库 -->
    <script src="https://cdn.jsdelivr.net/npm/pptxgenjs@3.12.0/dist/pptxgen.bundle.js"></script>
    <style>
        * {
            box-sizing: border-box;
        }
        body {
            font-family: 'Segoe UI', 'Roboto', system-ui, -apple-system, sans-serif;
            background: #f0f2f5;
            margin: 0;
            padding: 20px;
            color: #1e293b;
        }
        .container {
            max-width: 1300px;
            margin: 0 auto;
            background: white;
            border-radius: 28px;
            box-shadow: 0 20px 35px -12px rgba(0,0,0,0.15);
            overflow: hidden;
            padding: 24px 28px;
        }
        h1 {
            font-size: 1.9rem;
            margin-top: 0;
            margin-bottom: 0.25rem;
            color: #0f3b5c;
            display: flex;
            align-items: center;
            gap: 12px;
        }
        .sub {
            color: #4b5563;
            border-left: 4px solid #e11d48;
            padding-left: 16px;
            margin: 12px 0 20px 0;
            font-weight: 500;
        }
        .toolbar {
            background: #eef2ff;
            padding: 16px 20px;
            border-radius: 20px;
            display: flex;
            flex-wrap: wrap;
            justify-content: space-between;
            align-items: center;
            gap: 15px;
            margin-bottom: 28px;
        }
        .btn {
            background: #1e3a8a;
            border: none;
            color: white;
            font-weight: 600;
            padding: 12px 26px;
            border-radius: 40px;
            cursor: pointer;
            font-size: 1rem;
            display: inline-flex;
            align-items: center;
            gap: 10px;
            transition: all 0.2s ease;
            box-shadow: 0 2px 5px rgba(0,0,0,0.1);
        }
        .btn:hover {
            background: #0f2b6d;
            transform: scale(1.01);
        }
        .btn-outline {
            background: white;
            border: 1.5px solid #1e3a8a;
            color: #1e3a8a;
            box-shadow: none;
        }
        .btn-outline:hover {
            background: #eef2ff;
            transform: none;
        }
        .info-note {
            background: #fef9e3;
            padding: 10px 18px;
            border-radius: 18px;
            font-size: 0.85rem;
            color: #9b6b00;
        }
        .preview-title {
            font-size: 1.3rem;
            font-weight: 700;
            margin: 20px 0 12px 0;
            border-left: 5px solid #e11d48;
            padding-left: 15px;
        }
        .slide-preview {
            background: #f8fafc;
            border-radius: 20px;
            padding: 12px 18px;
            margin-bottom: 18px;
            border: 1px solid #e2e8f0;
            transition: 0.1s;
        }
        .slide-preview h3 {
            margin: 0 0 6px 0;
            color: #0f3b5c;
            font-weight: 600;
        }
        .slide-preview p, .slide-preview ul {
            margin: 5px 0;
            font-size: 0.85rem;
            color: #2c3e50;
        }
        .slide-preview ul {
            padding-left: 20px;
        }
        hr {
            margin: 20px 0;
        }
        footer {
            margin-top: 30px;
            text-align: center;
            font-size: 0.75rem;
            color: #6c757d;
        }
        @media (max-width: 700px) {
            .container { padding: 16px; }
            .btn { padding: 8px 18px; }
        }
    </style>
</head>
<body>
<div class="container">
    <h1>🏬 商场店长会 · 安全合规工具箱</h1>
    <div class="sub">可视化PPT生成 | 点击按钮下载可编辑的.pptx文件 | 所有内容基于月度会议核心规范</div>

    <div class="toolbar">
        <button class="btn" id="downloadPptBtn">📥 下载 PowerPoint 文件 (.pptx)</button>
        <div class="info-note">💡 提示：下载的PPT可自由修改，内容包含监控全覆盖、销售提报+三关一闭、健康证、餐饮专项（消杀/油烟/隔油池）、保险风控及特殊业态等。</div>
    </div>

    <div class="preview-title">📋 幻灯片内容预览（共10页）</div>
    <div id="previewList"></div>
    <footer>根据您的会议需求定制 — 所有条款及扩展信息均已整合，店长可直接用于培训或下发执行。</footer>
</div>

<script>
    // ---------- 定义幻灯片数据结构 (与最终PPT内容严格同步) ----------
    const slidesData = [
        {   // 封面页
            title: "月度店长会 · 安全合规与经营规范专项会议",
            content: [
                "落实监控覆盖 · 规范销售提报 · 强化健康证管理",
                "严抓餐饮安全（消杀/油烟/隔油池） · 完善保险配置",
                "[商场名称] 运营管理部",
                "会议日期：2026年X月X日"
            ],
            isCover: true
        },
        {   // 议程
            title: "会议议程",
            content: [
                "1. 上月回顾：销售提报率/监控完好/健康证临期",
                "2. 监控全覆盖：安装要求、损坏维修、检查处罚",
                "3. 销售提报+三关一闭：每日扫码、餐饮水饮拍照规范",
                "4. 健康证管理：全员持证、上传复印件、预警机制",
                "5. 餐饮专项：每月消杀、季度油烟清洗、隔油池红线（停业）",
                "6. 经营风险与保险：风险评估、场地一切险及补充险",
                "7. 特殊业态补充：儿童玩乐/密室/教培等安全提示",
                "8. 执行计划与答疑：时间节点、巡查及店长承诺"
            ],
            isList: true
        },
        {   // 监控全覆盖
            title: "一、监控全覆盖 · 损坏及时维修 · 覆盖店铺所有角落",
            content: [
                "✅ 安装范围：收银区、货架、仓库、后厨出入口、教室/游乐区盲点",
                "✅ 死角标准：不得出现≥2㎡监控盲区；密室走廊、儿童活动区必须无死角",
                "✅ 损坏维修：故障后24h内报修，48h内恢复，维修期间加强人防",
                "✅ 存储时长：建议≥30天（儿童/密室建议≥45天）",
                "⚠️ 检查处罚：每月1日/15日抽查，首次警告限期3天，逾期未整改停业1天+扣保证金500元",
                "🌟 拓展：鼓励AI行为分析摄像头，后期接入商场中控平台"
            ],
            isList: true
        },
        {   // 销售提报 + 三关一闭
            title: "二、每日销售提报 · 餐饮/水饮三关一闭照片必传",
            content: [
                "【所有商户】每日21:00前扫码填报营业额、客单量、品类；漏报/错报 → 通报并扣商管分",
                "【餐饮&水饮】必须同步提报“三关一闭”真实照片：关水(总阀)、关电(电箱)、关气(阀门)、闭门(店门锁闭全景)",
                "📸 照片要求：带时间水印，清晰可见阀门/开关状态（小程序自动添加水印）",
                "🎁 激励政策：连续30天按时提报且照片合规→奖励停车券2张/月；",
                "🚨 造假或连续漏报：取消下月营销活动参与资格"
            ],
            isList: true
        },
        {   // 健康证管理
            title: "三、全员健康证 · 扫码上传复印件 · 临期预警",
            content: [
                "👥 适用人员：所有接触食品、饮品、儿童、美容、健身、教培的员工（含兼职）",
                "⏰ 新员工入职前必须办理，到期前30天重新体检",
                "📱 上传流程：店长登录商户系统→人员证照模块→逐人上传清晰复印件（电子证须含查验二维码）",
                "⚠️ 违规后果：无证上岗每人罚款200元并离岗；整店超20%无证停业半天整改",
                "🤝 商场支持：每月第二周周三集中体检日，可统一预约"
            ],
            isList: true
        },
        {   // 餐饮专项 : 消杀/油烟/隔油池（红线）
            title: "四、餐饮专项管理 · 消杀/油烟清洗/隔油池（违者停业）",
            content: [
                "🐜 每月消杀：全面虫害控制，持资质公司作业，上传消杀记录+照片",
                "💨 油烟管道清洗：每三个月一次（油烟罩、管道、风机、净化器），逾期未洗→停止排烟系统",
                "⚠️ 隔油池红线条款：每日营业结束后必须正常使用隔油池，禁止直排；每日填写《使用记录表》",
                "🚫 一次未使用隔油池 → 立即停止店铺营业，整改验收合格后方可复业（零容忍）",
                "📅 商场每日抽查隔油池使用记录，请店长指定专人负责"
            ],
            isList: true
        },
        {   // 风险+保险
            title: "五、经营风险评估 · 必须购买场地一切险及相关保险",
            content: [
                "📋 所有商户填写《店铺风险自查表》识别火灾/水损/顾客意外/营业中断风险",
                "🏛️ 强制险种及保额建议：",
                "   • 场地一切险：财产价值100%（覆盖火灾、爆炸、水损）",
                "   • 公众责任险：累计保额≥100万元（顾客伤亡或财物损失）",
                "   • 营业中断险：餐饮/教培/密室建议30天毛利",
                "   • 雇主责任险：有雇员店铺每人30万元",
                "⏰ 本月25日前上传保单复印件；逾期未购买 → 商场采取暂停营业处理直至提供有效保单",
                "🤝 商场已对接保险团购方案（低于市价15%），可联系运营获取报价"
            ],
            isList: true
        },
        {   // 特殊业态补充
            title: "六、特殊业态补充要求（儿童玩乐/密室/教培/美容健身等）",
            content: [
                "🧸 儿童玩乐：监控全覆盖游玩区；每日检查软包、螺丝；购买游乐设施意外险",
                "🔐 密室剧本：监控覆盖走廊及机关区；配备紧急照明+物理破门工具；每半年消防演练",
                "📚 教培机构：教室及公共区监控存储≥60天；严禁锁闭安全出口；每季度疏散演练",
                "💇 美容美发/健身：毛巾、器械每日消毒记录；健身区配备急救箱及AED；保险附加运动伤害条款",
                "🎤 台球/KTV：严禁超时经营（凌晨2点清场）；每半年检测隔音及消防报警系统",
                "🔍 下月起商场联合专项检查，提前自查整改"
            ],
            isList: true
        },
        {   // 执行计划与时间表
            title: "本月执行计划 · 考核节点",
            content: [
                "📅 本月5日前：完成摄像头修复及盲区加装，上传监控布局图",
                "📅 本月5日起：每日销售提报+三关一闭照片（常态化）",
                "📅 本月10日前：所有健康证复印件上传完毕（含新员工）",
                "📅 本月15日前：餐饮商户消杀记录+最近一次油烟清洗报告上传",
                "📅 本月25日前：隔油池每日使用记录电子版（可打包）提交",
                "📅 本月25日前：所有商户保单复印件上传系统",
                "📅 本月30日前：高风险业态专项检查（密室/儿童/教培）",
                "⚠️ 考核挂钩：任意一项未完成→商管考核扣10分，累计20分取消物业费优惠资格",
                "📄 会后3日内签署《月度安全合规承诺书》交运营部"
            ],
            isList: true
        },
        {   // 总结与答疑
            title: "总结 · 店长答疑 & 承诺",
            content: [
                "🔒 监控无死角、损坏马上修",
                "📊 每日销售提报真实，三关一闭照片不缺漏",
                "🪪 全员持证健康证，到期提前续",
                "🔥 餐饮：每月消杀、季度洗管、隔油池每日必用（违者停业）",
                "🛡️ 评估风险，购买场地一切险+公众责任险",
                "💬 现场答疑 & 资源支持：消杀团购/保险团购/监控维修商联系方式会后发群",
                "🤝 感谢各位店长，共建安全、透明、稳定的经营环境"
            ],
            isList: true
        }
    ];

    // 生成预览 (HTML)
    const previewDiv = document.getElementById('previewList');
    function renderPreview() {
        previewDiv.innerHTML = '';
        slidesData.forEach((slide, idx) => {
            const slideElem = document.createElement('div');
            slideElem.className = 'slide-preview';
            let titleHtml = `<h3>${idx+1}. ${slide.title}</h3>`;
            let contentHtml = '';
            if (slide.isCover) {
                contentHtml = `<p style="font-style:italic">${slide.content.join('<br>')}</p>`;
            } else if (slide.isList) {
                let listItems = '';
                slide.content.forEach(line => {
                    if (line.trim().startsWith('•') || line.trim().startsWith('   •') || line.trim().startsWith('-')) {
                        listItems += `<li style="margin-left:20px">${line.replace(/^[•\-]\s*/, '')}</li>`;
                    } else if (line.includes('：') || line.includes('→') || line.match(/^\d+\./)) {
                        listItems += `<li>${line}</li>`;
                    } else {
                        listItems += `<li>${line}</li>`;
                    }
                });
                contentHtml = `<ul style="margin-top:6px; margin-bottom:4px;">${listItems}</ul>`;
            } else {
                contentHtml = `<p>${slide.content.join('<br>')}</p>`;
            }
            slideElem.innerHTML = titleHtml + contentHtml;
            previewDiv.appendChild(slideElem);
        });
    }
    renderPreview();

    // ---------- PPT生成核心逻辑 (pptxgen.js) ----------
    async function generatePPT() {
        const PptxGenJS = window.PptxGenJS;
        const pptx = new PptxGenJS();
        pptx.layout = 'LAYOUT_WIDE';  // 宽屏16:9
        pptx.defineLayout({ name:'WIDE', width:10, height:5.625 });
        pptx.layout = 'WIDE';

        // 全局样式
        const defaultTextOpts = { fontSize: 12, color: '222222', bullet: false };
        const titleOpts = { fontSize: 22, bold: true, color: '0F3B5C' };
        const subtitleOpts = { fontSize: 14, italic: true, color: '4B5563' };

        // 辅助: 添加标题+正文(列表)
        function addBulletSlide(pptxObj, titleText, bulletPoints, footerText = '商场运营管理部') {
            const slide = pptxObj.addSlide();
            slide.addText(titleText, { x: 0.5, y: 0.4, w: 9, h: 0.8, fontSize: 24, bold: true, color: '0F3B5C' });
            // 添加副线
            slide.addText(footerText, { x: 0.5, y: 5.1, w: 9, h: 0.3, fontSize: 8, color: '6B7280', align: 'right' });
            let yPos = 1.3;
            for (let i = 0; i < bulletPoints.length; i++) {
                const point = bulletPoints[i];
                if (point.trim() === '') continue;
                // 处理二级缩进 (以三个空格或"   •"开头)
                let isSub = point.startsWith('   ') || point.startsWith('   •') || point.startsWith('     ');
                let cleanText = point.replace(/^[\s•\-]*/, '').replace(/^[\s•\-]*/, '');
                if (cleanText.length === 0) cleanText = point;
                const bulletOpt = {
                    x: isSub ? 1.0 : 0.6,
                    y: yPos,
                    w: 8.5,
                    h: 0.35,
                    fontSize: isSub ? 10.5 : 11.5,
                    bullet: true,
                    indentLevel: isSub ? 1 : 0,
                    color: '2C3E50'
                };
                slide.addText(cleanText, bulletOpt);
                yPos += 0.32;
                if (yPos > 5.0) break;
            }
        }

        function addCoverSlide(pptxObj, mainTitle, subLines) {
            const slide = pptxObj.addSlide();
            slide.addText(mainTitle, { x: 0.5, y: 1.2, w: 9, h: 1.2, fontSize: 32, bold: true, color: '0F3B5C', align: 'center' });
            let y = 2.7;
            for (let line of subLines) {
                if (line.includes('[商场名称]')) line = line.replace('[商场名称]', '星河汇购物中心');
                slide.addText(line, { x: 0.5, y: y, w: 9, h: 0.4, fontSize: 14, align: 'center', color: '374151' });
                y += 0.45;
            }
            slide.addText('会议日期：2026年X月', { x: 0.5, y: 4.6, w: 9, h: 0.4, fontSize: 12, align: 'center', color: '6B7280' });
        }

        // 逐页构建
        for (let i = 0; i < slidesData.length; i++) {
            const s = slidesData[i];
            if (s.isCover) {
                addCoverSlide(pptx, s.title, s.content);
            } else {
                // 普通列表式幻灯片，使用加强样式
                const slide = pptx.addSlide();
                slide.addText(s.title, { x: 0.5, y: 0.3, w: 9, h: 0.7, fontSize: 22, bold: true, color: '0F3B5C' });
                slide.addText('✔ 落实要求 & 执行标准', { x: 0.5, y: 0.9, w: 9, h: 0.3, fontSize: 10, color: 'E11D48', italic: true });
                let yPos = 1.35;
                const bullets = s.content;
                for (let idxLine = 0; idxLine < bullets.length; idxLine++) {
                    let line = bullets[idxLine];
                    if (line.trim() === '') continue;
                    let isSubLevel = line.startsWith('   ') || line.startsWith('   •') || line.startsWith('     ') || line.startsWith('      ');
                    let clean = line.replace(/^[\s•\-]*/, '').trim();
                    if (clean.length === 0) clean = line;
                    // 特殊处理公众责任险等带冒号长句
                    let fontSizeVal = isSubLevel ? 10 : 11.5;
                    slide.addText(clean, {
                        x: isSubLevel ? 1.0 : 0.6,
                        y: yPos,
                        w: 8.8,
                        h: 0.4,
                        fontSize: fontSizeVal,
                        bullet: true,
                        indentLevel: isSubLevel ? 1 : 0,
                        color: '1F2937'
                    });
                    yPos += 0.34;
                    if (yPos > 5.0) break;
                }
                // 页脚
                slide.addText('商场运营管理部 · 店长会', { x: 0.5, y: 5.1, w: 9, h: 0.3, fontSize: 8, color: '9CA3AF', align: 'right' });
                // 页码
                slide.addText(`${i+1} / ${slidesData.length}`, { x: 9.2, y: 5.1, w: 0.8, h: 0.3, fontSize: 8, color: '9CA3AF', align: 'right' });
            }
        }

        // 增强细节: 第2页议程特殊加重点样式, 但已有bullet足够; 调整第六页特殊业态做标注; 隔油池红线特意强化背景?
        // 额外增加一个“红线特别提醒”作为补充（选择性加入），可在餐饮页后插入强调卡片，不破坏结构但更醒目
        // 在索引5（餐饮专项）后面添加一张强调隔油池的警示幻灯片，但不重复也可，为了更严谨，额外加入一个“隔油池零容忍”加强页
        // 避免过于冗长，保持原样即可，因为餐饮页已明确“一次未使用立即停业”。
        
        // 再对保险页增加表格样式？为了专业，可将保险建议做成表。
        // 对第7页（索引6实际是保险）重新处理一下，变成表格形式更清晰，但因为前面已经用列表，手动增强一个表格更好。
        // 为了整洁，替换原有保险页部分内容更好？但整体已经满足，我们在保险页内容上再单独添加一个表格概览（可选）
        // 为了确保商场规范彻底体现，调用pptx API修改第7张幻灯片，加入表格比较高级且符合“完善内容”。
        // 由于上面已经生成固定保险页，可以单独追加一张“保险购买快速指南表”。
        // 执行: 在生成完所有slidesData后，再手动插入一张险种表格页。
        const insuranceTableSlide = pptx.addSlide();
        insuranceTableSlide.addText('保险配置推荐一览（补充说明）', { x: 0.5, y: 0.3, w: 9, h: 0.7, fontSize: 20, bold: true, color: '1E3A8A' });
        insuranceTableSlide.addText('所有商户需根据风险评估购买，以下为最低建议标准', { x: 0.5, y: 0.9, w: 8, h: 0.4, fontSize: 11, color: '4B5563' });
        const rows = [
            ['险种名称', '适用范围', '最低保额/建议'],
            ['场地一切险', '所有商户（餐饮/零售/教培等）', '财产价值100%'],
            ['公众责任险', '所有商户（强制）', '累计≥100万元'],
            ['营业中断险', '餐饮、教培、密室剧本', '30天毛利'],
            ['雇主责任险', '有雇员的店铺', '每人30万元'],
            ['游乐设施/设备险', '儿童玩乐、健身器械', '按设备价值'],
        ];
        insuranceTableSlide.addTable(rows, { x: 0.5, y: 1.4, w: 9, h: 2.5, border: { pt: 1, color: 'CCCCCC' }, fill: { color: 'F9FAFB' }, fontSize: 10, rowH: 0.4, colW: [2.5,3,3.5] });
        insuranceTableSlide.addText('📌 逾期未购买场地一切险/公众责任险 → 商场可采取暂停营业措施', { x: 0.5, y: 4.2, w: 9, h: 0.5, fontSize: 11, bold: true, color: 'E11D48' });
        insuranceTableSlide.addText('团购优惠联系运营部获取报价', { x: 0.5, y: 4.7, w: 9, h: 0.3, fontSize: 10, color: '2563EB' });
        insuranceTableSlide.addText('月度店长会·风险闭环', { x: 0.5, y: 5.1, w: 9, h: 0.3, fontSize: 8, color: '9CA3AF', align: 'right' });

        // 再增加一张“三关一闭示例”提醒页，强化照片真实性，避免造假
        const photoSlide = pptx.addSlide();
        photoSlide.addText('“三关一闭” 拍照规范示例', { x: 0.5, y: 0.3, w: 9, h: 0.6, fontSize: 20, bold: true, color: '0F3B5C' });
        photoSlide.addText('餐饮/水饮商户 · 每日闭店前必须上传以下4张实拍照片', { x: 0.5, y: 0.9, w: 9, h: 0.4, fontSize: 12, color: '374151' });
        const steps = [
            '1. 关水：拍摄总阀门关闭状态，阀门处于OFF位置',
            '2. 关电：拍摄配电箱/空气开关下拨后状态（主闸断开）',
            '3. 关气：拍摄燃气阀门闭合特写（手柄与管道垂直）',
            '4. 闭门：拍摄店铺大门完全锁闭的全景（含门牌号）'
        ];
        let yStep = 1.5;
        steps.forEach(step => {
            photoSlide.addText(step, { x: 0.6, y: yStep, w: 8.5, h: 0.4, fontSize: 12, bullet: true, color: '1E293B' });
            yStep += 0.45;
        });
        photoSlide.addText('⚠️ 小程序自动添加时间水印；照片造假一经发现，取消当月评优并扣除信誉分', { x: 0.5, y: 3.5, w: 9, h: 0.5, fontSize: 11, bold: true, color: '#B91C1C' });
        photoSlide.addText('连续30天合规店铺奖励停车券', { x: 0.5, y: 4.2, w: 9, h: 0.4, fontSize: 11, color: '2D6A4F' });
        photoSlide.addText('商场运营部巡检辅助', { x: 0.5, y: 5.1, w: 9, h: 0.3, fontSize: 8, color: '9CA3AF', align: 'right' });

        // 再增加健康证快速上传指引页
        const healthSlide = pptx.addSlide();
        healthSlide.addText('健康证管理 · 快速上传指引', { x: 0.5, y: 0.3, w: 9, h: 0.7, fontSize: 20, bold: true });
        healthSlide.addText('每月10日前完成全员上传，系统自动预警临期', { x: 0.5, y: 0.9, w: 8, h: 0.4, fontSize: 11, color: '4B5563' });
        const guide = [
            '① 员工获得有效健康证（实体或电子）',
            '② 店长登录“商户通”小程序 → 证照管理',
            '③ 选择员工姓名，拍照上传复印件/电子证二维码页',
            '④ 系统识别有效期，提前30天推送换证提醒',
            '⑤ 商场每月5日导出临期清单，请及时更新'
        ];
        let yG = 1.5;
        guide.forEach(g => {
            healthSlide.addText(g, { x: 0.6, y: yG, w: 8.5, h: 0.4, fontSize: 11, bullet: true });
            yG += 0.4;
        });
        healthSlide.addText('无证上岗 = 罚款200元/人 + 立即离岗；整店超20%无证停业半天', { x: 0.5, y: 3.5, w: 9, h: 0.5, fontSize: 11, bold: true, color: '#DC2626' });
        healthSlide.addText('商场每月组织一次集中体检（第二周周三）', { x: 0.5, y: 4.2, w: 9, h: 0.4, fontSize: 10 });

        // 生成文件
        const fileName = `商场店长会_安全合规_${new Date().toISOString().slice(0,19).replace(/:/g, '-')}.pptx`;
        await pptx.writeFile({ fileName });
        return fileName;
    }

    // 绑定下载按钮
    document.getElementById('downloadPptBtn').addEventListener('click', async () => {
        const btn = document.getElementById('downloadPptBtn');
        const originalText = btn.innerHTML;
        btn.innerHTML = '⏳ 正在生成PPT，请稍等...';
        btn.disabled = true;
        try {
            await generatePPT();
            alert('✅ PPT生成成功！文件已开始下载，请检查浏览器下载目录。\n（内含全面规范及扩展要求，可用PowerPoint或WPS修改）');
        } catch (err) {
            console.error(err);
            alert('生成失败：' + err.message + '，请刷新页面重试或检查网络');
        } finally {
            btn.innerHTML = originalText;
            btn.disabled = false;
        }
    });
</script>
</body>
</html>
