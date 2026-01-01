
import React, { useState, useMemo, useEffect } from 'react';
import { v4 as uuidv4 } from 'uuid';
import { 
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, 
  WidthType, AlignmentType, VerticalAlign 
} from 'docx';
import XLSX_Lib from 'xlsx';
import { ActivityData, ScheduleItem, ExpenseItem } from './types';

const XLSX: any = (XLSX_Lib as any).utils ? XLSX_Lib : (XLSX_Lib as any).default || XLSX_Lib;

// 字体与字号配置
const FONT_SONG = "仿宋";
const FONT_HEI = "黑体";
const SIZE_TITLE = 36;
const SIZE_CONTENT = 32;
const SIZE_TABLE = 24;

const getWeekday = (dateStr: string) => {
  const date = new Date(dateStr);
  const weekdays = ['日', '一', '二', '三', '四', '五', '六'];
  return weekdays[date.getDay()];
};

const triggerDownload = (blob: Blob, filename: string) => {
  try {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  } catch (e) {
    console.error("下载失败", e);
    alert("下载失败，请尝试在电脑端使用 Chrome 浏览器。");
  }
};

const initialSchedule: ScheduleItem[] = [
  { id: uuidv4(), time: '10:00-10:30', content: '签到入场', speaker: '李润林' },
  { id: uuidv4(), time: '10:30-11:00', content: '公司介绍', speaker: '李润林' },
  { id: uuidv4(), time: '11:00-12:00', content: '业务专题培训', speaker: '李润林' },
  { id: uuidv4(), time: '12:00-13:30', content: '午餐及休息', speaker: '' },
  { id: uuidv4(), time: '13:30-15:00', content: '产品方案宣导', speaker: '李润林' },
  { id: uuidv4(), time: '15:00-16:30', content: '研讨及通关', speaker: '渠道团队长' },
];

const initialExpenses = (): ExpenseItem[] => [
  { id: uuidv4(), category: '住宿费', project: '仅限于封闭培训期间发生的住宿费', price: 0, unit: '间/晚', quantity: 0, total: 0, description: '' },
  { id: uuidv4(), category: '交通费', project: '交通费、租车费', price: 0, unit: '项', quantity: 0, total: 0, description: '' },
  { id: uuidv4(), category: '餐费', project: '培训期间的正餐餐费', price: 0, unit: '元/人/天', quantity: 0, total: 0, description: '' },
  { id: uuidv4(), category: '茶点费', project: '培训期间的茶点费', price: 0, unit: '元/人/天', quantity: 0, total: 0, description: '' },
  { id: uuidv4(), category: '场地、设备租赁费', project: '培训场地、培训专用设备的租赁费', price: 0, unit: '场', quantity: 0, total: 0, description: '' },
  { id: uuidv4(), category: '培训资料、文具费', project: '印制讲师、学员手册、相关培训书籍、资料、文具等费用', price: 0, unit: '项', quantity: 0, total: 0, description: '' },
  { id: uuidv4(), category: '外聘教师课时费', project: '聘请公司外讲师进行培训授课的课时费', price: 0, unit: '课时', quantity: 0, total: 0, description: '' },
  { id: uuidv4(), category: '培训活动费', project: '仅限于七天以上培训用于观摩、考察等费用', price: 0, unit: '项', quantity: 0, total: 0, description: '' },
  { id: uuidv4(), category: '培训宣传费', project: '提升培训效果宣传用品费(横幅、展板、胸卡等)', price: 0, unit: '项', quantity: 0, total: 0, description: '' },
  { id: uuidv4(), category: '其他费用', project: '学员合影留念的照片、培训现场照片等制作费用', price: 0, unit: '项', quantity: 0, total: 0, description: '' },
];

const App: React.FC = () => {
  const [isLoggedIn, setIsLoggedIn] = useState(false);
  const [loginLoading, setLoginLoading] = useState(false);
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');

  const [data, setData] = useState<ActivityData>(() => {
    const saved = localStorage.getItem('activity_form_v20');
    if (saved) {
      try { return JSON.parse(saved); } catch (e) { console.error(e); }
    }
    return {
      channelName: '大童保险销售服务有限公司四川分公司',
      activityDate: new Date().toISOString().split('T')[0],
      startTime: '09:30',
      endTime: '16:30',
      location: '成都市锦江区东大路段',
      participantsDesc: '大童川分部分绩优人员、复保工作人员',
      submitDate: new Date().toISOString().split('T')[0],
      schedule: initialSchedule,
      participantCount: 0,
      expenses: initialExpenses(),
    };
  });

  useEffect(() => {
    localStorage.setItem('activity_form_v20', JSON.stringify(data));
  }, [data]);

  const handleLogin = (e: React.FormEvent) => {
    e.preventDefault();
    setLoginLoading(true);
    // 模拟登录延迟
    setTimeout(() => {
      setIsLoggedIn(true);
      setLoginLoading(false);
    }, 800);
  };

  const handleInputChange = (e: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement>) => {
    const { name, value } = e.target;
    setData(prev => {
      const next = { ...prev, [name]: value };
      if (name === 'channelName') {
        const shortName = value.substring(0, 2);
        next.participantsDesc = `${shortName}川分部分绩优人员、复保工作人员`;
      }
      return next;
    });
  };

  const updateScheduleItem = (id: string, field: keyof ScheduleItem, value: string) => {
    setData(prev => ({
      ...prev,
      schedule: prev.schedule.map(item => item.id === id ? { ...item, [field]: value } : item)
    }));
  };

  const updateSpecificExpense = (project: string, field: keyof ExpenseItem, value: any) => {
    setData(prev => ({
      ...prev,
      expenses: prev.expenses.map(exp => {
        if (exp.project === project) {
          const updated = { ...exp, [field]: value };
          if (field === 'price' || field === 'quantity') {
            const p = parseFloat(updated.price as any) || 0;
            const q = parseFloat(updated.quantity as any) || 0;
            updated.total = p * q;
          }
          return updated;
        }
        return exp;
      })
    }));
  };

  const getExpenseByProject = (project: string) => {
    return data.expenses.find(e => e.project === project) || { price: 0, unit: '', quantity: 0, total: 0, description: '' };
  };

  const totalExpense = useMemo(() => data.expenses.reduce((sum, item) => sum + (Number(item.total) || 0), 0), [data.expenses]);

  const generateWord = async () => {
    const [y, m, d] = data.activityDate.split('-');
    const [sy, sm, sd] = data.submitDate.split('-');
    const dateStr = data.activityDate.replace(/-/g, '');
    const weekday = getWeekday(data.activityDate);
    
    const doc = new Document({
      sections: [{
        properties: {},
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({ 
                text: `关于举办复星保德信四川分公司与${data.channelName}团队的培训通知`, 
                bold: true, 
                size: SIZE_TITLE, 
                font: FONT_HEI 
              }),
            ],
            spacing: { after: 400 },
          }),
          new Paragraph({
            children: [
              new TextRun({ 
                text: `根据四川分公司中介条线发展规划，分公司定于 ${y} 年 ${m} 月 ${d} 日（星期${weekday}）举办与${data.channelName}团队的培训。具体安排如下：`, 
                size: SIZE_CONTENT, 
                font: FONT_SONG 
              }),
            ],
            indent: { firstLine: 480 },
            spacing: { line: 400 },
          }),
          new Paragraph({
            children: [
              new TextRun({ text: `一、活动时间：`, bold: true, size: SIZE_CONTENT, font: FONT_SONG }),
              new TextRun({ text: `${y}年${m}月${d}日 ${data.startTime}至${data.endTime}`, size: SIZE_CONTENT, font: FONT_SONG })
            ],
            spacing: { before: 200, line: 400 },
          }),
          new Paragraph({
            children: [
              new TextRun({ text: `二、活动地点：`, bold: true, size: SIZE_CONTENT, font: FONT_SONG }),
              new TextRun({ text: data.location, size: SIZE_CONTENT, font: FONT_SONG })
            ],
            spacing: { line: 400 },
          }),
          new Paragraph({
            children: [
              new TextRun({ text: `三、参加人员：`, bold: true, size: SIZE_CONTENT, font: FONT_SONG }),
              new TextRun({ text: data.participantsDesc, size: SIZE_CONTENT, font: FONT_SONG })
            ],
            spacing: { line: 400 },
          }),
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "时间", font: FONT_SONG, size: SIZE_TABLE })], alignment: AlignmentType.CENTER })], verticalAlign: VerticalAlign.CENTER }),
                  new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "内容", font: FONT_SONG, size: SIZE_TABLE })], alignment: AlignmentType.CENTER })], verticalAlign: VerticalAlign.CENTER }),
                  new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "主讲人", font: FONT_SONG, size: SIZE_TABLE })], alignment: AlignmentType.CENTER })], verticalAlign: VerticalAlign.CENTER }),
                ],
              }),
              ...data.schedule.map(item => new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: item.time, font: FONT_SONG, size: SIZE_TABLE })], alignment: AlignmentType.CENTER })] }),
                  new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: item.content, font: FONT_SONG, size: SIZE_TABLE })], alignment: AlignmentType.CENTER })] }),
                  new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: item.speaker, font: FONT_SONG, size: SIZE_TABLE })], alignment: AlignmentType.CENTER })] }),
                ],
              })),
            ],
          }),
          new Paragraph({
            children: [new TextRun({ text: "复星保德信四川分公司", size: SIZE_CONTENT, font: FONT_SONG })],
            alignment: AlignmentType.RIGHT,
            spacing: { before: 800 },
          }),
          new Paragraph({
            children: [new TextRun({ text: `${sy}年${sm}月${sd}日`, size: SIZE_CONTENT, font: FONT_SONG })],
            alignment: AlignmentType.RIGHT,
          }),
        ],
      }],
    });

    const blob = await Packer.toBlob(doc);
    triggerDownload(blob, `${dateStr}${data.channelName}培训活动通知.docx`);
  };

  const generateExcel = () => {
    try {
      if (!XLSX || !XLSX.utils) throw new Error("XLSX 库加载失败");

      const borderThin = {
        top: { style: 'thin' }, bottom: { style: 'thin' },
        left: { style: 'thin' }, right: { style: 'thin' }
      };
      
      const styleTitle = {
        font: { name: '宋体', sz: 14, bold: true },
        alignment: { horizontal: 'center', vertical: 'center' }
      };

      const styleHeader = {
        font: { name: '宋体', sz: 11, color: { rgb: "FFFFFF" }, bold: true },
        fill: { fgColor: { rgb: "376092" } },
        alignment: { horizontal: 'center', vertical: 'center' },
        border: borderThin
      };

      const styleCenter = {
        font: { name: '宋体', sz: 11 },
        alignment: { horizontal: 'center', vertical: 'center' },
        border: borderThin
      };

      const styleCenterWrapped = {
        font: { name: '宋体', sz: 11 },
        alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
        border: borderThin
      };

      const styleLeftWrapped = {
        font: { name: '宋体', sz: 11 },
        alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
        border: borderThin
      };

      const styleNoBorder = { border: {} };

      const wb = XLSX.utils.book_new();
      const ws: any = { '!ref': 'A1:G30' };

      // A1:G1 标题 (无边框)
      ws['A1'] = { v: '费用明细-培训类', s: styleTitle };
      for(let c=1; c<=6; c++) ws[XLSX.utils.encode_cell({r: 0, c: c})] = { v: "", s: styleNoBorder };

      // A2:G2 无边框
      for(let c=0; c<=6; c++) ws[XLSX.utils.encode_cell({r: 1, c: c})] = { v: "", s: styleNoBorder };

      // A3:B3 培训举办地
      ws['A3'] = { v: '培训举办地', s: styleCenter };
      ws['B3'] = { v: '', s: styleCenter };
      ws['C3'] = { v: data.location, s: styleLeftWrapped };
      ws['D3'] = { v: '', s: styleLeftWrapped };
      ws['E3'] = { v: '', s: styleLeftWrapped };

      // A4:B4 预估参与人数
      ws['A4'] = { v: '预估参与人数', s: styleCenter };
      ws['B4'] = { v: '', s: styleCenter };
      ws['C4'] = { v: data.participantCount || "", s: styleCenter };
      ws['D4'] = { v: '', s: styleCenter };
      ws['E4'] = { v: '', s: styleCenter };

      // 第五行无边框
      for(let c=0; c<=6; c++) ws[XLSX.utils.encode_cell({r: 4, c: c})] = { v: "", s: styleNoBorder };

      // 第六行：表头
      const headers = ["项目", "费用项目", "单价", "单位", "数量", "总价", "费用说明"];
      headers.forEach((h, i) => {
        ws[XLSX.utils.encode_cell({ r: 5, c: i })] = { v: h, s: styleHeader };
      });

      const setRow = (r: number, cat: string, proj: string, p: any, u: string, q: any, t: any, d: string) => {
        const isNumericBlank = (!t || Number(t) === 0);
        ws[XLSX.utils.encode_cell({ r, c: 0 })] = { v: cat, s: styleCenterWrapped }; 
        ws[XLSX.utils.encode_cell({ r, c: 1 })] = { v: proj, s: styleLeftWrapped };   
        ws[XLSX.utils.encode_cell({ r, c: 2 })] = { v: isNumericBlank ? "" : p, s: styleCenter };
        ws[XLSX.utils.encode_cell({ r, c: 3 })] = { v: isNumericBlank ? "" : u, s: styleCenter };
        ws[XLSX.utils.encode_cell({ r, c: 4 })] = { v: isNumericBlank ? "" : q, s: styleCenter };
        ws[XLSX.utils.encode_cell({ r, c: 5 })] = { v: isNumericBlank ? "" : t, s: styleCenter };
        ws[XLSX.utils.encode_cell({ r, c: 6 })] = { v: isNumericBlank ? "" : d, s: styleLeftWrapped };
      };

      const getE = (proj: string) => data.expenses.find(e => e.project === proj) || { category: '', project: '', price: 0, unit: '', quantity: 0, total: 0, description: '' };
      
      const expsConfig = [
        { row: 6, proj: '仅限于封闭培训期间发生的住宿费', cat: '住宿费' },
        { row: 7, proj: '交通费、租车费', cat: '交通费' },
        { row: 8, proj: '教务人员大型培训用具搬运费', cat: '' },
        { row: 9, proj: '培训期间的正餐餐费', cat: '餐费' },
        { row: 10, proj: '培训期间的茶点费', cat: '' },
        { row: 11, proj: '培训场地、培训专用设备的租赁费', cat: '场地、设备租赁费' },
        { row: 12, proj: '印制讲师、学员手册、相关培训书籍、资料、文具等费用', cat: '培训资料、文具费' },
        { row: 13, proj: '聘请公司外讲师进行培训授课的课时费', cat: '外聘教师课时费' },
        { row: 14, proj: '仅限于七天以上培训用于观摩、考察等费用', cat: '培训活动费' },
        { row: 15, proj: '提升培训效果宣传用品费(横幅、展板、胸卡等)', cat: '培训宣传费' },
        { row: 16, proj: '学员合影留念的照片、培训现场照片等制作费用', cat: '其他费用' },
        { row: 17, proj: '常用药品购买费用', cat: '' },
        { row: 18, proj: '教务公杂费', cat: '' }
      ];

      expsConfig.forEach(conf => {
        const e = getE(conf.proj);
        const description = (conf.proj === '培训期间的正餐餐费') ? '' : e.description;
        setRow(conf.row, conf.cat, conf.proj, e.price, e.unit, e.quantity, e.total, description);
      });

      ws['A20'] = { v: '合计', s: styleCenter };
      ws['B20'] = { v: '', s: styleCenter };
      ws['F20'] = { v: totalExpense || "", s: styleCenter };
      ['C20', 'D20', 'E20', 'G20'].forEach(cell => ws[cell] = {v: "", s: styleCenter});

      for(let r=6; r<=19; r++){
        for(let c=0; c<=6; c++){
          const addr = XLSX.utils.encode_cell({r, c});
          if(!ws[addr]) ws[addr] = { v: "", s: styleCenter };
        }
      }

      for(let r=20; r<=21; r++){
        for(let c=0; c<=6; c++) ws[XLSX.utils.encode_cell({r, c})] = { v: "", s: styleNoBorder };
      }

      ws['A23'] = { v: '备注：', s: { font: { name: '宋体', sz: 11 } } };
      ws['B23'] = { v: '1.以上费用均为含税价（包含但不限于税费、服务费等）；', s: { font: { name: '宋体', sz: 11 } } };
      ws['B24'] = { v: '2.模板供参考，可按实际情况调整；', s: { font: { name: '宋体', sz: 11 } } };
      ws['B25'] = { v: '3.除费用明细表外，还需提供会议/培训通知、会议/培训日程安排等支持材料；', s: { font: { name: '宋体', sz: 11 } } };

      ws['!merges'] = [
        { s: { r: 0, c: 0 }, e: { r: 0, c: 6 } }, 
        { s: { r: 2, c: 0 }, e: { r: 2, c: 1 } }, 
        { s: { r: 2, c: 2 }, e: { r: 2, c: 4 } }, 
        { s: { r: 3, c: 0 }, e: { r: 3, c: 1 } }, 
        { s: { r: 3, c: 2 }, e: { r: 3, c: 4 } }, 
        { s: { r: 7, c: 0 }, e: { r: 8, c: 0 } }, 
        { s: { r: 9, c: 0 }, e: { r: 10, c: 0 } }, 
        { s: { r: 16, c: 0 }, e: { r: 18, c: 0 } }, 
        { s: { r: 19, c: 0 }, e: { r: 19, c: 1 } }, 
      ];

      ws['!cols'] = [{ wch: 10.82 }, { wch: 22.45 }, { wch: 8 }, { wch: 10.55 }, { wch: 23.09 }, { wch: 8 }, { wch: 28.73 }];

      XLSX.utils.book_append_sheet(wb, ws, "费用明细");
      const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      triggerDownload(blob, `${data.activityDate.replace(/-/g, '')}${data.channelName}培训费用明细表.xlsx`);

    } catch (err: any) {
      console.error("Excel 生成失败:", err);
      alert("Excel 生成失败: " + err.message);
    }
  };

  // Define lunchExp and teaExp derived from data to fix "Cannot find name" errors in JSX
  const lunchExp = getExpenseByProject('培训期间的正餐餐费');
  const teaExp = getExpenseByProject('培训期间的茶点费');

  if (!isLoggedIn) {
    return (
      <div className="min-h-screen bg-slate-900 flex items-center justify-center p-6 relative overflow-hidden">
        {/* 背景装饰 */}
        <div className="absolute top-0 left-0 w-full h-full">
            <div className="absolute top-[-10%] left-[-10%] w-[40%] h-[40%] bg-blue-600/20 rounded-full blur-[120px]"></div>
            <div className="absolute bottom-[-10%] right-[-10%] w-[40%] h-[40%] bg-indigo-600/20 rounded-full blur-[120px]"></div>
        </div>

        <div className="w-full max-w-[420px] relative z-10 animate-in fade-in slide-in-from-bottom-8 duration-700">
          <div className="bg-white/95 backdrop-blur-xl rounded-[2.5rem] shadow-2xl p-10 border border-white/20">
            <div className="text-center mb-10">
              <div className="inline-flex items-center justify-center w-16 h-16 bg-blue-600 rounded-2xl mb-6 shadow-xl shadow-blue-500/20">
                <i className="fas fa-file-signature text-white text-2xl"></i>
              </div>
              <h1 className="text-2xl font-black text-slate-900 tracking-tight">签报管理系统</h1>
              <p className="text-slate-400 text-sm mt-2 font-medium">Fosun United Health - Sichuan</p>
            </div>

            <form onSubmit={handleLogin} className="space-y-6">
              <div className="space-y-4">
                <div className="relative group">
                  <span className="absolute left-5 top-1/2 -translate-y-1/2 text-slate-300 group-focus-within:text-blue-500 transition-colors">
                    <i className="fas fa-user-circle"></i>
                  </span>
                  <input 
                    type="text" 
                    placeholder="用户名" 
                    value={username}
                    onChange={(e) => setUsername(e.target.value)}
                    className="w-full bg-slate-50 border border-slate-100 rounded-2xl py-4 pl-12 pr-6 outline-none focus:border-blue-500 focus:bg-white transition-all text-sm font-semibold"
                  />
                </div>
                <div className="relative group">
                  <span className="absolute left-5 top-1/2 -translate-y-1/2 text-slate-300 group-focus-within:text-blue-500 transition-colors">
                    <i className="fas fa-lock"></i>
                  </span>
                  <input 
                    type="password" 
                    placeholder="访问密码" 
                    value={password}
                    onChange={(e) => setPassword(e.target.value)}
                    className="w-full bg-slate-50 border border-slate-100 rounded-2xl py-4 pl-12 pr-6 outline-none focus:border-blue-500 focus:bg-white transition-all text-sm font-semibold"
                  />
                </div>
              </div>

              <button 
                type="submit" 
                disabled={loginLoading}
                className="w-full bg-blue-600 hover:bg-blue-700 disabled:bg-blue-400 text-white py-4 rounded-2xl font-black text-sm shadow-xl shadow-blue-600/20 active:scale-[0.98] transition-all flex items-center justify-center gap-3"
              >
                {loginLoading ? (
                  <i className="fas fa-circle-notch animate-spin"></i>
                ) : (
                  <>进入系统 <i className="fas fa-arrow-right text-[10px]"></i></>
                )}
              </button>
            </form>

            <div className="mt-10 text-center">
              <p className="text-[10px] text-slate-300 uppercase font-black tracking-widest">内部办公专用 · 严禁泄露</p>
            </div>
          </div>
          <p className="text-center text-slate-500 text-[11px] mt-8 font-medium">© 2024 复星保德信四川分公司 中介条线</p>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-50 font-sans text-slate-900 pb-20">
      <header className="bg-white border-b border-slate-200 px-8 py-5 flex justify-between items-center sticky top-0 z-40 shadow-sm">
        <div className="flex items-center gap-3">
          <div className="bg-blue-600 p-2 rounded-lg shadow-lg shadow-blue-600/20"><i className="fas fa-file-signature text-white text-xl"></i></div>
          <div>
             <h1 className="text-xl font-bold tracking-tight leading-none">签报助手 <span className="text-blue-600">v5.0</span></h1>
             <span className="text-[10px] text-slate-400 font-bold uppercase tracking-tighter">Activity Report Management</span>
          </div>
        </div>
        <div className="flex items-center gap-6">
          <div className="hidden md:flex flex-col items-end mr-2">
            <span className="text-[11px] font-black text-slate-400 uppercase leading-none">当前用户</span>
            <span className="text-xs font-bold text-slate-600">管理员</span>
          </div>
          <button onClick={() => { if(confirm('重置将清空所有内容？')){ localStorage.clear(); window.location.reload(); } }} className="px-4 py-2 text-sm font-semibold text-slate-400 hover:text-red-500 transition-colors">重置</button>
          <button onClick={() => setIsLoggedIn(false)} className="px-4 py-2 text-sm font-semibold text-slate-400 hover:text-slate-600 transition-colors">退出</button>
          <button onClick={() => { generateWord(); setTimeout(generateExcel, 800); }} className="bg-blue-600 hover:bg-blue-700 text-white px-6 py-2.5 rounded-xl font-bold text-sm shadow-lg active:scale-95 transition-all">一键导出</button>
        </div>
      </header>

      <main className="max-w-7xl mx-auto p-8 grid grid-cols-1 lg:grid-cols-3 gap-8 animate-in fade-in duration-500">
        <div className="lg:col-span-2 space-y-8">
          <section className="bg-white rounded-3xl p-8 shadow-sm border border-slate-100">
            <h2 className="text-lg font-bold mb-6 flex items-center"><span className="w-1.5 h-6 bg-yellow-400 rounded-full mr-3"></span>基础信息录入</h2>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              <div className="md:col-span-2">
                <label className="block text-xs font-black text-slate-400 uppercase mb-2">渠道名称 (全称)</label>
                <input name="channelName" value={data.channelName} onChange={handleInputChange} className="w-full px-5 py-4 bg-yellow-50 border-2 border-yellow-100 rounded-2xl focus:bg-white focus:border-blue-500 outline-none font-bold text-lg transition-all" />
              </div>
              <div className="md:col-span-2 grid grid-cols-1 md:grid-cols-3 gap-4">
                 <div>
                    <label className="block text-xs font-black text-slate-400 uppercase mb-2">培训日期</label>
                    <input type="date" name="activityDate" value={data.activityDate} onChange={handleInputChange} className="w-full px-5 py-4 bg-yellow-50 border-2 border-yellow-100 rounded-2xl font-bold outline-none" />
                 </div>
                 <div>
                    <label className="block text-xs font-black text-slate-400 uppercase mb-2">开始时间</label>
                    <input type="time" name="startTime" value={data.startTime} onChange={handleInputChange} className="w-full px-5 py-4 bg-yellow-50 border-2 border-yellow-100 rounded-2xl font-bold outline-none" />
                 </div>
                 <div>
                    <label className="block text-xs font-black text-slate-400 uppercase mb-2">结束时间</label>
                    <input type="time" name="endTime" value={data.endTime} onChange={handleInputChange} className="w-full px-5 py-4 bg-yellow-50 border-2 border-yellow-100 rounded-2xl font-bold outline-none" />
                 </div>
              </div>
              <div>
                <label className="block text-xs font-black text-slate-400 uppercase mb-2">培训地点</label>
                <input name="location" value={data.location} onChange={handleInputChange} className="w-full px-5 py-4 bg-yellow-50 border-2 border-yellow-100 rounded-2xl font-bold outline-none" />
              </div>
              <div className="md:col-span-2">
                <label className="block text-xs font-black text-slate-400 uppercase mb-2">参加人员 (自动简称/可修改)</label>
                <input name="participantsDesc" value={data.participantsDesc} onChange={handleInputChange} className="w-full px-5 py-4 bg-yellow-50 border-2 border-yellow-100 rounded-2xl font-bold outline-none" />
              </div>
            </div>
          </section>

          <section className="bg-white rounded-3xl p-8 shadow-sm border border-slate-100">
            <h2 className="text-lg font-bold mb-6 flex items-center"><span className="w-1.5 h-6 bg-blue-500 rounded-full mr-3"></span>日程流程安排</h2>
            <div className="overflow-hidden border border-slate-100 rounded-2xl">
              <table className="w-full text-left">
                <thead className="bg-slate-50"><tr className="text-xs font-bold text-slate-400 border-b border-slate-100"><th className="px-6 py-4">时间</th><th className="px-6 py-4">内容</th><th className="px-6 py-4">主讲人</th></tr></thead>
                <tbody className="divide-y divide-slate-50">
                  {data.schedule.map(item => (
                    <tr key={item.id} className="text-sm">
                      <td className="px-6 py-3 font-semibold text-slate-600"><input value={item.time} onChange={(e) => updateScheduleItem(item.id, 'time', e.target.value)} className="w-full border-none p-0 focus:ring-0 outline-none" /></td>
                      <td className="px-6 py-3 text-slate-600"><input value={item.content} onChange={(e) => updateScheduleItem(item.id, 'content', e.target.value)} className="w-full border-none p-0 focus:ring-0 outline-none" /></td>
                      <td className="px-6 py-3 text-slate-600"><input value={item.speaker} onChange={(e) => updateScheduleItem(item.id, 'speaker', e.target.value)} className="w-full border-none p-0 focus:ring-0 outline-none" /></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </section>
        </div>

        <aside className="space-y-8">
          {/* 餐饮录入 */}
          <div className="bg-white rounded-3xl p-6 border border-slate-100 shadow-sm space-y-4">
             <h2 className="text-sm font-bold flex items-center gap-2"><i className="fas fa-utensils text-orange-400"></i> 餐饮费录入</h2>
             <div className="space-y-4">
                <div className="p-4 bg-orange-50/50 rounded-2xl border border-orange-100">
                   <p className="text-[10px] font-black text-orange-400 uppercase mb-3 tracking-tighter">正餐餐费</p>
                   <div className="grid grid-cols-2 gap-2 mb-2">
                      <div className="space-y-1">
                        <span className="text-[9px] text-slate-400 font-bold ml-1">单价</span>
                        <input placeholder="0" type="number" value={lunchExp.price || ''} onChange={(e) => updateSpecificExpense('培训期间的正餐餐费', 'price', e.target.value)} className="w-full bg-white border border-slate-200 rounded-lg px-3 py-2 text-xs font-bold focus:border-orange-300 outline-none transition-all" />
                      </div>
                      <div className="space-y-1">
                        <span className="text-[9px] text-slate-400 font-bold ml-1">单位</span>
                        <input placeholder="元/人/天" value={lunchExp.unit || ''} onChange={(e) => updateSpecificExpense('培训期间的正餐餐费', 'unit', e.target.value)} className="w-full bg-white border border-slate-200 rounded-lg px-3 py-2 text-xs font-bold focus:border-orange-300 outline-none transition-all" />
                      </div>
                   </div>
                   <div className="grid grid-cols-2 gap-2">
                      <div className="space-y-1">
                        <span className="text-[9px] text-slate-400 font-bold ml-1">数量</span>
                        <input placeholder="0" type="number" value={lunchExp.quantity || ''} onChange={(e) => updateSpecificExpense('培训期间的正餐餐费', 'quantity', e.target.value)} className="w-full bg-white border border-slate-200 rounded-lg px-3 py-2 text-xs font-bold focus:border-orange-300 outline-none transition-all" />
                      </div>
                      <div className="space-y-1">
                        <span className="text-[9px] text-slate-400 font-bold ml-1">总价</span>
                        <div className="flex items-center bg-orange-100 border border-orange-200 rounded-lg px-3 py-2 text-xs font-black text-orange-600 min-h-[32px]">￥{lunchExp.total}</div>
                      </div>
                   </div>
                </div>

                <div className="p-4 bg-blue-50/50 rounded-2xl border border-blue-100">
                   <p className="text-[10px] font-black text-blue-400 uppercase mb-3 tracking-tighter">茶点费</p>
                   <div className="grid grid-cols-2 gap-2 mb-2">
                      <div className="space-y-1">
                        <span className="text-[9px] text-slate-400 font-bold ml-1">单价</span>
                        <input placeholder="0" type="number" value={teaExp.price || ''} onChange={(e) => updateSpecificExpense('培训期间的茶点费', 'price', e.target.value)} className="w-full bg-white border border-slate-200 rounded-lg px-3 py-2 text-xs font-bold focus:border-blue-300 outline-none transition-all" />
                      </div>
                      <div className="space-y-1">
                        <span className="text-[9px] text-slate-400 font-bold ml-1">单位</span>
                        <input placeholder="元/人/天" value={teaExp.unit || ''} onChange={(e) => updateSpecificExpense('培训期间的茶点费', 'unit', e.target.value)} className="w-full bg-white border border-slate-200 rounded-lg px-3 py-2 text-xs font-bold focus:border-blue-300 outline-none transition-all" />
                      </div>
                   </div>
                   <div className="grid grid-cols-2 gap-2">
                      <div className="space-y-1">
                        <span className="text-[9px] text-slate-400 font-bold ml-1">数量</span>
                        <input placeholder="0" type="number" value={teaExp.quantity || ''} onChange={(e) => updateSpecificExpense('培训期间的茶点费', 'quantity', e.target.value)} className="w-full bg-white border border-slate-200 rounded-lg px-3 py-2 text-xs font-bold focus:border-blue-300 outline-none transition-all" />
                      </div>
                      <div className="space-y-1">
                        <span className="text-[9px] text-slate-400 font-bold ml-1">总价</span>
                        <div className="flex items-center bg-blue-100 border border-blue-200 rounded-lg px-3 py-2 text-xs font-black text-blue-600 min-h-[32px]">￥{teaExp.total}</div>
                      </div>
                   </div>
                </div>
             </div>
          </div>

          {/* 统计 */}
          <div className="bg-slate-900 rounded-[2.5rem] p-8 text-white shadow-2xl relative overflow-hidden">
             <div className="absolute -top-10 -right-10 w-40 h-40 bg-blue-500/10 rounded-full blur-3xl"></div>
             <h2 className="text-sm font-black text-blue-400 uppercase mb-6 tracking-widest">预算统计</h2>
             <div className="space-y-6">
                <div>
                   <label className="block text-[10px] text-slate-400 mb-2">预估总人数 (Excel A4单元格用)</label>
                   <div className="flex items-center"><input type="number" name="participantCount" value={data.participantCount || ''} onChange={handleInputChange} className="bg-transparent text-5xl font-black outline-none w-32 border-b border-white/10" /><span className="text-xl ml-3 opacity-40 font-bold">人</span></div>
                </div>
                <div>
                   <label className="block text-[10px] text-slate-400 mb-2 font-black uppercase">费用预算合计</label>
                   <div className="text-4xl font-black text-green-400 font-mono tracking-tighter">¥ {totalExpense.toLocaleString()}</div>
                </div>
                <div className="pt-6 border-t border-white/10">
                   <label className="block text-[10px] text-slate-400 mb-3">落款日期</label>
                   <input type="date" name="submitDate" value={data.submitDate} onChange={handleInputChange} className="w-full bg-white/5 border border-white/10 rounded-xl px-4 py-3 text-sm font-bold outline-none focus:bg-white/10 transition-all" />
                </div>
             </div>
          </div>
          
          <div className="grid grid-cols-1 gap-4">
             <button onClick={generateWord} className="w-full bg-white border border-slate-200 py-4 rounded-2xl font-bold text-slate-700 hover:bg-slate-50 transition-all flex items-center justify-center gap-2 group"><i className="fas fa-file-word text-blue-500 group-hover:scale-110 transition-transform"></i> 导出培训通知 (Word)</button>
             <button onClick={generateExcel} className="w-full bg-white border border-slate-200 py-4 rounded-2xl font-bold text-slate-700 hover:bg-slate-50 transition-all flex items-center justify-center gap-2 group"><i className="fas fa-file-excel text-green-600 group-hover:scale-110 transition-transform"></i> 导出费用明细 (Excel)</button>
          </div>
        </aside>
      </main>
    </div>
  );
};

export default App;
