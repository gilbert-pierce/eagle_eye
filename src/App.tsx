import React, { useState, useRef, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import { 
  Shield, 
  FileUp, 
  FileDown, 
  AlertTriangle, 
  CheckCircle2, 
  Search,
  Database,
  Info,
  Filter,
  RotateCcw,
  Trash2,
  Table,
  Play,
  RefreshCw
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { cn } from './lib/utils';
import { ReimbursementRecord, AuditRisk } from './types';


interface LeavePeriod {
  employee: string;
  start: string;
  end: string;
  type: string;
}

interface AuditConfig {
  enabledRules: {
    ghostHost: boolean;
    frequency: boolean;
    venues: boolean;
    perCapita: boolean;
    drinking: boolean;
    soe: boolean;
    leave: boolean;
    topNEmployee: boolean;
    topNDepartment: boolean;
    travelOverlap: boolean;
    travelLocationMismatch: boolean;
    splitTravel: boolean;
    invalidGuest: boolean;
    internalGroup: boolean;
  };
  ghostHost: {
    hosts: string[];
    targets: string[];
  };
  frequency: {
    limit: number;
    consecutiveDays: number;
    holidays: string[];
  };
  venues: {
    keywords: string[];
  };
  perCapita: {
    limit: number;
    ratio: number;
  };
  drinking: {
    noKeywords: string[];
    alcoholKeywords: string[];
  };
  soe: {
    keywords: string[];
  };
  leave: {
    periods: LeavePeriod[];
  };
  topN: {
    n: number;
    thresholdMultiplier: number;
  };
  invalidGuest: {
    keywords: string[];
  };
  internalGroup: {
    keywords: string[];
  };
  fieldMappings: {
    employee: string;
    department: string;
    amount: string;
    date: string;
    host: string;
    guest: string;
    description: string;
    alcoholStatus: string;
    alcoholAmount: string;
    perCapita: string;
    guestCount: string;
    staffCount: string;
    receptionType: string;
    travelDays: string;
    orderId: string;
    isRemote: string;
    expenseType: string;
    travelReason: string;
    staffNames: string;
    organization: string;
  };
}

const KNOWN_HOLIDAYS = [
  // 2024
  '2024-01-01', '2024-02-10', '2024-02-11', '2024-02-12', '2024-02-13', '2024-02-14', '2024-02-15', '2024-02-16', '2024-02-17',
  '2024-04-04', '2024-04-05', '2024-04-06', '2024-05-01', '2024-05-02', '2024-05-03', '2024-05-04', '2024-05-05',
  '2024-06-10', '2024-09-15', '2024-09-16', '2024-09-17', '2024-10-01', '2024-10-02', '2024-10-03', '2024-10-04', '2024-10-05', '2024-10-06', '2024-10-07',
  // 2025
  '2025-01-01', '2025-01-28', '2025-01-29', '2025-01-30', '2025-01-31', '2025-02-01', '2025-02-02', '2025-02-03', '2025-02-04',
  '2025-04-04', '2025-04-05', '2025-04-06', '2025-05-01', '2025-05-02', '2025-05-03', '2025-05-04', '2025-05-05',
  '2025-05-31', '2025-06-01', '2025-06-02', '2025-10-01', '2025-10-02', '2025-10-03', '2025-10-04', '2025-10-05', '2025-10-06', '2025-10-07', '2025-10-08',
  // 2026
  '2026-01-01', '2026-02-16', '2026-02-17', '2026-02-18', '2026-02-19', '2026-02-20', '2026-02-21', '2026-02-22',
  '2026-04-04', '2026-04-05', '2026-04-06', '2026-05-01', '2026-05-02', '2026-05-03', '2026-05-04', '2026-05-05',
  '2026-06-20', '2026-06-21', '2026-06-22', '2026-09-25', '2026-09-26', '2026-09-27', '2026-10-01', '2026-10-02', '2026-10-03', '2026-10-04', '2026-10-05', '2026-10-06', '2026-10-07'
];

const RULE_DEFINITIONS = [
  { 
    key: 'ghostHost', 
    name: '虚假主持人', 
    description: '检测主持人是否为司机、助理等定义角色，且招待对象为敏感实体。',
    example: '举例：主持人字段为“王五（司机）”，招待对象字段包含“技术供应商”。默认使用字段：[应酬招待主持人]、[招待对象]。',
    defaultFields: ['应酬招待主持人', '招待对象'],
    mappingKeys: ['host', 'guest'],
    params: [
      { key: 'hosts', label: '主持人关键词', type: 'text' },
      { key: 'targets', label: '敏感对象关键词', type: 'text' }
    ]
  },
  { 
    key: 'frequency', 
    name: '报销频次', 
    description: '检测同一员工在同一天提交的报销单数量是否超过设定上限，或连续n日发生招待（自动跳过周末及配置的法定节假日）。公式：count(报销单) > 每日报销单上限，或 连续招待天数 >= 连续报销天数上限。',
    example: '举例：员工“张三”在“2026-03-20”提交了3笔报销，超过上限2笔；或周四、周五、下周一连续发生招待（跨周末视作连续）。默认使用字段：[费用所属员工]、[发生日期]。',
    defaultFields: ['费用所属员工', '发生日期'],
    mappingKeys: ['employee', 'date'],
    params: [
      { key: 'limit', label: '每日报销单上限 (笔)', type: 'number' },
      { key: 'consecutiveDays', label: '连续报销天数上限', type: 'number' },
      { key: 'holidays', label: '法定节假日', type: 'textarea' }
    ]
  },
  { 
    key: 'venues', 
    name: '违规场所', 
    description: '检测费用描述中是否包含KTV、会所、按摩等违规场所关键词。',
    example: '举例：费用描述包含“KTV”、“按摩”、“洗浴”等。默认使用字段：[费用描述（招待原因）]。',
    defaultFields: ['费用描述（招待原因）'],
    mappingKeys: ['description'],
    params: [
      { key: 'keywords', label: '违规场所关键词', type: 'textarea' }
    ]
  },
  { 
    key: 'perCapita', 
    name: '人均/陪餐', 
    description: '检测招待人均金额是否超标，以及行内陪餐人数比例是否合规。公式：招待人均金额 > 人均限额 或 行内陪餐人员人数 > 招待对象人数 * 陪餐比例。',
    example: '举例：人均金额超过500元，或陪餐人数超过招待对象人数。默认使用字段：[招待人均金额]、[行内陪餐人员人数]、[招待对象人数]。',
    defaultFields: ['招待人均金额', '行内陪餐人员人数', '招待对象人数'],
    mappingKeys: ['perCapita', 'staffCount', 'guestCount'],
    params: [
      { key: 'limit', label: '人均限额 (元)', type: 'number' },
      { key: 'ratio', label: '陪餐比例 (陪:客)', type: 'number' }
    ]
  },
  { 
    key: 'drinking', 
    name: '酒水异常', 
    description: '检测是否存在提及酒水但未登记的情况。',
    example: '举例：描述提及“茅台”但未登记。默认使用字段：[是否领用行内酒水]、[费用描述（招待原因）]。',
    defaultFields: ['是否领用行内酒水', '费用描述（招待原因）'],
    mappingKeys: ['alcoholStatus', 'description'],
    params: [
      { key: 'alcoholKeywords', label: '酒水关键词', type: 'text' },
      { key: 'noKeywords', label: '未领用标记', type: 'text' }
    ]
  },
  { 
    key: 'soe', 
    name: '国企总部', 
    description: '检测招待对象疑似国企总部人员时，接待类型是否违规设置为商务接待。',
    example: '举例：招待对象包含“央企”、“总部”，但接待类型为“商务接待”而非“公务接待”。默认使用字段：[招待对象]、[接待类型]。',
    defaultFields: ['招待对象', '接待类型'],
    mappingKeys: ['guest', 'receptionType'],
    params: [
      { key: 'keywords', label: '总部/央企关键词', type: 'text' }
    ]
  },
  { 
    key: 'leave', 
    name: '休假期间', 
    description: '检测员工在年假、病假等休假期间是否存在本地报销记录。',
    example: '举例：员工在年假期间产生“非异地”报销。默认使用字段：[费用所属员工]、[发生日期]、[是否异地]。',
    defaultFields: ['费用所属员工', '发生日期', '是否异地'],
    mappingKeys: ['employee', 'date', 'isRemote'],
    params: [
      { key: 'periods', label: '异常期间列表', type: 'textarea' }
    ]
  },
  { 
    key: 'topNEmployee', 
    name: '个人TopN', 
    description: '分析累计报销金额排名前N的员工，检测其是否显著高于平均水平。公式：该员工总额 > 其他员工平均总额 * 异常倍数阈值。',
    example: '举例：前5名员工报销总额超过其他员工平均水平的3倍。默认使用字段：[费用所属员工]、[折人民币金额]。',
    defaultFields: ['费用所属员工', '折人民币金额'],
    mappingKeys: ['employee', 'amount'],
    params: [
      { key: 'n', label: '分析前 N 名', type: 'number' },
      { key: 'thresholdMultiplier', label: '异常倍数阈值', type: 'number' }
    ]
  },
  { 
    key: 'topNDepartment', 
    name: '部门TopN', 
    description: '分析累计报销金额排名前N的部门，检测其是否显著高于平均水平。公式：该部门总额 > 其他部门平均总额 * 异常倍数阈值。',
    example: '举例：前5名部门报销总额超过其他部门平均水平的3倍。默认使用字段：[部门]、[折人民币金额]。',
    defaultFields: ['部门', '折人民币金额'],
    mappingKeys: ['department', 'amount'],
    params: [
      { key: 'n', label: '分析前 N 名', type: 'number' },
      { key: 'thresholdMultiplier', label: '异常倍数阈值', type: 'number' }
    ]
  },
  { 
    key: 'travelOverlap', 
    name: '差旅重复', 
    description: '检测同一员工名下是否存在出差时间重叠的报销记录。公式：max(开始日期1, 开始日期2) < min(结束日期1, 结束日期2)。',
    example: '举例：两笔差旅单据的时间段存在交叉。默认使用字段：[费用所属员工]、[发生日期]、[出差天数]。',
    defaultFields: ['费用所属员工', '发生日期', '出差天数'],
    mappingKeys: ['employee', 'date', 'travelDays'],
    params: []
  },
  { 
    key: 'travelLocationMismatch', 
    name: '出差地点矛盾', 
    description: '检测员工申请异地出差时是否存在本地招待记录，或反之。',
    example: '举例：申请了异地出差，但同日有本地招待记录（作为主持人或陪餐人员）。默认使用字段：[费用所属员工]、[发生日期]、[是否异地]、[费用类型]、[应酬招待主持人]、[行内陪餐人员]。',
    defaultFields: ['费用所属员工', '发生日期', '是否异地', '费用类型', '应酬招待主持人', '行内陪餐人员'],
    mappingKeys: ['employee', 'date', 'isRemote', 'expenseType', 'host', 'staffNames'],
    params: []
  },
  { 
    key: 'splitTravel', 
    name: '事由分拆', 
    description: '检测同一员工针对同一出差事由是否存在多笔分拆报销的行为。',
    example: '举例：同一出差事由“某项目调研”对应多笔报销单。默认使用字段：[费用所属员工]、[出差事由]、[费用类型]。',
    defaultFields: ['费用所属员工', '出差事由', '费用类型'],
    mappingKeys: ['employee', 'travelReason', 'expenseType'],
    params: []
  },
  { 
    key: 'invalidGuest', 
    name: '对象异常', 
    description: '检测招待对象是否包含司机、师傅等不符合商业招待必要性的身份。',
    example: '举例：招待对象包含“司机”、“师傅”。默认使用字段：[招待对象]。',
    defaultFields: ['招待对象'],
    mappingKeys: ['guest'],
    params: [
      { key: 'keywords', label: '异常身份关键词', type: 'text' }
    ]
  },
  { 
    key: 'internalGroup', 
    name: '内部接待', 
    description: '检测是否发生了集团内部禁止的相互接待行为（涉及xib, cyb等关键词）。',
    example: '举例：招待对象包含“集团内部”、“xib”等。默认使用字段：[招待对象]、[费用描述（招待原因）]。',
    defaultFields: ['招待对象', '费用描述（招待原因）'],
    mappingKeys: ['guest', 'description'],
    params: [
      { key: 'keywords', label: '内部接待关键词', type: 'text' }
    ]
  },
];

const DEFAULT_CONFIG: AuditConfig = {
  enabledRules: {
    ghostHost: true,
    frequency: true,
    venues: true,
    perCapita: true,
    drinking: true,
    soe: true,
    leave: true,
    topNEmployee: true,
    topNDepartment: true,
    travelOverlap: true,
    travelLocationMismatch: true,
    splitTravel: true,
    invalidGuest: true,
    internalGroup: true,
  },
  ghostHost: {
    hosts: ['Driver', 'Assistant', '司机', '助理'],
    targets: ['Investment', 'Tech', 'Provider', '投资', '技术', '供应商'],
  },
  frequency: {
    limit: 2,
    consecutiveDays: 3,
    holidays: KNOWN_HOLIDAYS
  },
  venues: {
    keywords: ['KTV', 'Club', 'Spa', 'Massage', '会所', '洗浴', '按摩'],
  },
  perCapita: {
    limit: 500,
    ratio: 1, // 1:1
  },
  drinking: {
    noKeywords: ['No', '否', 'N'],
    alcoholKeywords: ['酒', '茅台', '五粮液', '剑南春', '红酒', '白酒', '洋酒'],
  },
  soe: {
    keywords: ['国有', '总部', '央企', '国企', 'State-Owned', 'HQ'],
  },
  leave: {
    periods: [
      { employee: '张三', start: '2026-03-01', end: '2026-03-10', type: '年假' },
      { employee: '李四', start: '2026-03-15', end: '2026-03-20', type: '出境考察' }
    ],
  },
  topN: {
    n: 5,
    thresholdMultiplier: 3,
  },
  invalidGuest: {
    keywords: ['司机', 'driver', '师傅', '保安', '家属'],
  },
  internalGroup: {
    keywords: ['xib', 'cyb', 'lib', 'xill', '集团内部', '内部接待'],
  },
  fieldMappings: {
    employee: '费用所属员工',
    department: '部门',
    amount: '折人民币金额',
    date: '发生日期',
    host: '应酬招待主持人',
    guest: '招待对象',
    description: '费用描述（招待原因）',
    alcoholStatus: '是否领用行内酒水',
    alcoholAmount: '酒水金额',
    perCapita: '招待人均金额',
    guestCount: '招待对象人数',
    staffCount: '行内陪餐人员人数',
    receptionType: '接待类型',
    travelDays: '出差天数',
    orderId: '报销单编号',
    isRemote: '是否异地',
    expenseType: '费用类型',
    travelReason: '出差事由',
    staffNames: '陪餐人员',
    organization: '机构',
  },
};

const MAPPING_LABELS: Record<string, string> = {
  employee: '费用所属员工',
  department: '部门',
  amount: '折人民币金额',
  date: '发生日期',
  host: '应酬招待主持人',
  guest: '招待对象',
  description: '费用描述（招待原因）',
  alcoholStatus: '是否领用行内酒水',
  alcoholAmount: '酒水金额',
  perCapita: '招待人均金额',
  guestCount: '招待对象人数',
  staffCount: '行内陪餐人员人数',
  receptionType: '接待类型',
  travelDays: '出差天数',
  orderId: '报销单编号',
  isRemote: '是否异地',
  expenseType: '费用类型',
  travelReason: '出差事由',
  staffNames: '陪餐人员',
  organization: '机构',
};

const DEFAULT_HEADERS = [
  '机构', '部门', '责任中心', '报销人', '费用所属员工', '收款人', '费用所属客户', '党组织/党务部门', '报销单编号', '费用类型', '费用科目', '费用描述（招待原因）', '币种', '原币金额', '折人民币金额', '反冲金额', '税额', '状态', '发生日期', '提交日期', '记账日期', '当前审批人', '已审批人', '审批日期', '审批信息', '附件张数', '费用终止日期', '出差天数', '旅差费开始结束日期', '应酬招待主持人', '主持人人数', '行内陪餐人员', '行内陪餐人员人数', '工作人员', '工作人员人数', '招待对象', '招待对象人数', '招待人均金额', '工作人员人均金额', '赠品名称', '接待类型', '是否领用行内酒水', '酒水金额', '是否异地', '出差地点', '契约号', '关联发票', 'HR系统审批单号', '特殊情况说明', '出差事由', '陪餐人员'
];

export default function App() {
  const [records, setRecords] = useState<ReimbursementRecord[]>([]);
  const [risks, setRisks] = useState<AuditRisk[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [isAuditOutdated, setIsAuditOutdated] = useState(false);
  const [fileName, setFileName] = useState<string | null>(null);
  const [toastMessage, setToastMessage] = useState<{message: string, type: 'success' | 'error'} | null>(null);

  const showToast = (message: string, type: 'success' | 'error' = 'success') => {
    setToastMessage({ message, type });
    setTimeout(() => setToastMessage(null), 3000);
  };
  const [config, setConfig] = useState<AuditConfig>(DEFAULT_CONFIG);
  const [headers, setHeaders] = useState<string[]>(DEFAULT_HEADERS);
  const [showConfig, setShowConfig] = useState(false);
  const [availableExpenseTypes, setAvailableExpenseTypes] = useState<string[]>([]);
  const [selectedExpenseTypes, setSelectedExpenseTypes] = useState<string[]>([]);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const leaveFileInputRef = useRef<HTMLInputElement>(null);

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setFileName(file.name);
    setIsProcessing(true);

    // Yield to main thread to allow loading UI to render
    setTimeout(() => {
      const reader = new FileReader();
      reader.onload = async (evt) => {
        try {
          const bstr = evt.target?.result;
          const wb = XLSX.read(bstr, { type: 'binary' });
          const wsname = wb.SheetNames[0];
          const ws = wb.Sheets[wsname];
          const data = XLSX.utils.sheet_to_json(ws) as any[];
        
        // Extract headers
        const sheetHeaders = XLSX.utils.sheet_to_json(ws, { header: 1 })[0] as string[];
        if (sheetHeaders && sheetHeaders.length > 0) {
          setHeaders(sheetHeaders);
        }

        // Map data to our struct
        const mappedRecords: ReimbursementRecord[] = data.map((row, index) => {
          const recordId = row['报销单编号'] || `REC-${index + 1000}`;
          return {
            id: recordId,
            机构: row['机构'] || '',
            部门: row['部门'] || '',
            责任中心: row['责任中心'] || '',
            报销人: row['报销人'] || '',
            费用所属员工: row['费用所属员工'] || '未知',
            收款人: row['收款人'] || '',
            费用所属客户: row['费用所属客户'] || '',
            '党组织/党务部门': row['党组织/党务部门'] || '',
            报销单编号: recordId,
            费用类型: row['费用类型'] || '',
            费用科目: row['费用科目'] || '',
            '费用描述（招待原因）': row['费用描述（招待原因）'] || '',
            币种: row['币种'] || 'CNY',
            原币金额: parseFloat(row['原币金额'] || 0),
            折人民币金额: parseFloat(row['折人民币金额'] || 0),
            反冲金额: parseFloat(row['反冲金额'] || 0),
            税额: parseFloat(row['税额'] || 0),
            状态: row['状态'] || '',
            发生日期: row['发生日期'] || '',
            提交日期: row['提交日期'] || '',
            记账日期: row['记账日期'] || '',
            当前审批人: row['当前审批人'] || '',
            已审批人: row['已审批人'] || '',
            审批日期: row['审批日期'] || '',
            审批信息: row['审批信息'] || '',
            附件张数: parseInt(row['附件张数'] || 0),
            费用终止日期: row['费用终止日期'] || '',
            出差天数: parseInt(row['出差天数'] || 0),
            旅差费开始结束日期: row['旅差费开始结束日期'] || '',
            应酬招待主持人: row['应酬招待主持人'] || '',
            主持人人数: parseInt(row['主持人人数'] || 0),
            行内陪餐人员: row['行内陪餐人员'] || '',
            行内陪餐人员人数: parseInt(row['行内陪餐人员人数'] || 0),
            工作人员: row['工作人员'] || '',
            工作人员人数: parseInt(row['工作人员人数'] || 0),
            招待对象: row['招待对象'] || '',
            招待对象人数: parseInt(row['招待对象人数'] || 0),
            招待人均金额: parseFloat(row['招待人均金额'] || 0),
            工作人员人均金额: parseFloat(row['工作人员人均金额'] || 0),
            赠品名称: row['赠品名称'] || '',
            接待类型: row['接待类型'] || '',
            是否领用行内酒水: row['是否领用行内酒水'] || '否',
            酒水金额: parseFloat(row['酒水金额'] || 0),
            是否异地: row['是否异地'] || '否',
            出差地点: row['出差地点'] || '',
            契约号: row['契约号'] || '',
            关联发票: row['关联发票'] || '',
            HR系统审批单号: row['HR系统审批单号'] || '',
            特殊情况说明: row['特殊情况说明'] || '',
            ...row
          };
        });

        setRecords(mappedRecords);
        
        // Extract unique expense types
        const types = Array.from(new Set(mappedRecords.map(r => r[config.fieldMappings.expenseType]).filter(Boolean)));
        setAvailableExpenseTypes(types);
        setSelectedExpenseTypes(types); // Default to all selected
        
        await runAudit(mappedRecords, config, types);
      } catch (error) {
        console.error("Error parsing Excel:", error);
        showToast("Failed to parse Excel file. Please ensure it is a valid .xlsx or .xls file.", 'error');
        setIsProcessing(false);
      }
    };
    reader.readAsBinaryString(file);
    }, 50);
  };

  const handleLeaveFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const data = XLSX.utils.sheet_to_json(ws) as any[];

        const periods: LeavePeriod[] = data.map(row => {
          // Try to find matching columns
          const employee = row['姓名'] || row['员工姓名'] || row['employee'] || row['name'] || '';
          
          // Handle Excel date numbers
          const parseDate = (val: any) => {
            if (!val) return '';
            if (typeof val === 'number') {
              const date = new Date((val - (25567 + 2)) * 86400 * 1000);
              return date.toISOString().split('T')[0];
            }
            return String(val);
          };
          
          const start = parseDate(row['开始日期'] || row['休假开始'] || row['start'] || '');
          const end = parseDate(row['结束日期'] || row['休假结束'] || row['end'] || '');
          const type = row['休假类型'] || row['类型'] || row['type'] || '休假';

          if (employee && start && end) {
            return { employee, start, end, type };
          }
          return null;
        }).filter(Boolean) as LeavePeriod[];

        if (periods.length > 0) {
          updateConfig({ leave: { periods } });
          showToast(`成功导入 ${periods.length} 条休假记录`);
        } else {
          showToast('未能从文件中识别到有效的休假记录，请确保包含：姓名、开始日期、结束日期等列。', 'error');
        }
      } catch (error) {
        console.error("Error parsing leave file:", error);
        showToast('解析文件失败，请检查文件格式。', 'error');
      }
    };
    reader.readAsBinaryString(file);
    // Reset input
    if (leaveFileInputRef.current) {
      leaveFileInputRef.current.value = '';
    }
  };

  const runAudit = (
    data: ReimbursementRecord[], 
    currentConfig: AuditConfig = config, 
    currentSelectedTypes: string[] = selectedExpenseTypes
  ): Promise<void> => {
    return new Promise((resolve) => {
      setIsProcessing(true);
      
      // Yield to main thread to allow loading UI to render
      setTimeout(() => {
        const detectedRisks: AuditRisk[] = [];
        const enabled = currentConfig.enabledRules;

        // Filter data by selected expense types
        const filteredData = data.filter(r => currentSelectedTypes.includes(String(r[currentConfig.fieldMappings.expenseType])));

    // Rule 2: Frequency Overload (Helper map)
    const frequencyMap: Record<string, number> = {};
    const employeeDates: Record<string, Set<string>> = {};
    
    // Top N Analysis Helpers
    const employeeTotals: Record<string, number> = {};
    const deptTotals: Record<string, number> = {};
    
    // Travel Overlap Helpers
    const travelIntervals: Record<string, { start: number, end: number, id: string, date: string }[]> = {};
    
    // Rule 1: Travel Location Mismatch Helpers
    const travelStatusByDate: Record<string, Record<string, string>> = {}; // employee -> date -> isRemote
    const receptionStatusByDate: Record<string, { isRemote: string, id: string, role: string, date: string }[]> = {}; // employee -> date -> status[]

    // Rule 2: Split Travel Helpers
    const travelReasonMap: Record<string, Set<string>> = {}; // `${employee}_${travelReason}` -> Set<orderId>

    filteredData.forEach(record => {
      const m = currentConfig.fieldMappings;
      const recordId = record[m.orderId] || record.id;
      const amount = parseFloat(record[m.amount]) || 0;
      const employee = String(record[m.employee] || '未知');
      const department = String(record[m.department] || '未知');
      const date = String(record[m.date] || '');
      const expenseType = String(record[m.expenseType] || '').toLowerCase();
      const isRemote = String(record[m.isRemote] || '');
      
      // Accumulate totals for Top N
      employeeTotals[employee] = (employeeTotals[employee] || 0) + amount;
      deptTotals[department] = (deptTotals[department] || 0) + amount;

      // Ghost Host Audit
      if (enabled.ghostHost) {
        const hostLower = String(record[m.host] || '').toLowerCase();
        const guestLower = String(record[m.guest] || '').toLowerCase();
        
        const isGhostHost = currentConfig.ghostHost.hosts.some(k => hostLower.includes(k.toLowerCase()));
        const isSensitiveTarget = currentConfig.ghostHost.targets.some(k => guestLower.includes(k.toLowerCase()));

        if (isGhostHost && isSensitiveTarget) {
          detectedRisks.push({
            recordId: recordId,
            ruleName: '虚假主持人',
            severity: 'high',
            description: `主持人 (${record[m.host]}) 为定义角色，但招待对象 (${record[m.guest]}) 为敏感实体。`
          });
        }
      }

      // Blacklisted Venues
      if (enabled.venues) {
        const descUpper = String(record[m.description] || '').toUpperCase();
        if (currentConfig.venues.keywords.some(k => descUpper.includes(k.toUpperCase()))) {
          detectedRisks.push({
            recordId: recordId,
            ruleName: '违规场所',
            severity: 'high',
            description: `费用描述中包含违规场所关键词。`
          });
        }
      }

      // Drinking Anomaly (Enhanced)
      if (enabled.drinking) {
        const descUpper = String(record[m.description] || '').toUpperCase();
        const noAlcohol = currentConfig.drinking.noKeywords;
        const isMarkedNo = noAlcohol.some(k => String(record[m.alcoholStatus]).toLowerCase().includes(k.toLowerCase()));
        const descMentionsAlcohol = currentConfig.drinking.alcoholKeywords.some(k => descUpper.includes(k.toUpperCase()));

        if (isMarkedNo && descMentionsAlcohol) {
          detectedRisks.push({
            recordId: recordId,
            ruleName: '酒水登记缺失',
            severity: 'medium',
            description: `费用描述提及酒水关键词，但未登记领用（标记为"${record[m.alcoholStatus]}"），疑似隐匿酒水费用。`
          });
        }
      }

      // SOE HQ Audit (New)
      if (enabled.soe) {
        const guestLower = String(record[m.guest] || '').toLowerCase();
        const isSOE = currentConfig.soe.keywords.some(k => guestLower.includes(k.toLowerCase()));
        if (isSOE && record[m.receptionType] === '商务接待') {
          detectedRisks.push({
            recordId: recordId,
            ruleName: '接待类型违规',
            severity: 'medium',
            description: `招待对象 (${record[m.guest]}) 疑似国企总部人员，但接待类型被列为 "商务接待" 而非 "公务接待" 。`
          });
        }
      }

      // Leave/Outbound Audit (New)
      if (enabled.leave) {
        const leaveMatch = currentConfig.leave.periods.find(p => 
          p.employee === employee && 
          date >= p.start && 
          date <= p.end
        );
        if (leaveMatch && String(record[m.isRemote]) === '否') {
          detectedRisks.push({
            recordId: recordId,
            ruleName: '休假期间本地报销',
            severity: 'high',
            description: `员工在 ${leaveMatch.type} 期间 (${leaveMatch.start} 至 ${leaveMatch.end}) 存在本地报销记录。`
          });
        }
      }

      // Per Capita Audit
      if (enabled.perCapita) {
        const perCapitaVal = parseFloat(record[m.perCapita]) || 0;
        const staffCount = parseInt(record[m.staffCount]) || 0;
        const guestCount = parseInt(record[m.guestCount]) || 0;

        if (perCapitaVal > currentConfig.perCapita.limit) {
          detectedRisks.push({
            recordId: recordId,
            ruleName: '人均超标',
            severity: 'medium',
            description: `招待人均金额 (${perCapitaVal}) 超过自定义限额 (${currentConfig.perCapita.limit})。`
          });
        }
        if (staffCount > guestCount * currentConfig.perCapita.ratio) {
          detectedRisks.push({
            recordId: recordId,
            ruleName: '陪餐人数超标',
            severity: 'low',
            description: `行内陪餐人员 (${staffCount}) 超过招待对象人数 (${guestCount}) 的 ${currentConfig.perCapita.ratio} 倍。`
          });
        }
      }

      // Frequency Overload Tracking
      if (enabled.frequency) {
        const freqKey = `${employee}_${date}`;
        frequencyMap[freqKey] = (frequencyMap[freqKey] || 0) + 1;
        
        if (!employeeDates[employee]) employeeDates[employee] = new Set();
        if (date) employeeDates[employee].add(date);
      }
      
      // Travel Overlap Tracking
      const travelDays = parseInt(record[m.travelDays]) || 0;
      if (enabled.travelOverlap && travelDays > 0) {
        const start = new Date(date).getTime();
        const end = start + (travelDays * 24 * 60 * 60 * 1000);
        if (!travelIntervals[employee]) travelIntervals[employee] = [];
        travelIntervals[employee].push({ start, end, id: recordId, date: date });
      }

      // Rule 1: Travel Location Mismatch Tracking
      if (enabled.travelLocationMismatch) {
        if (expenseType.includes('差旅')) {
          if (!travelStatusByDate[employee]) travelStatusByDate[employee] = {};
          travelStatusByDate[employee][date] = isRemote;
        }
        if (expenseType.includes('招待')) {
          const host = String(record[m.host] || '');
          const staff = String(record[m.staffNames] || '');
          
          // Check if current employee is host or staff
          if (host === employee || staff.includes(employee)) {
            if (!receptionStatusByDate[employee]) receptionStatusByDate[employee] = [];
            receptionStatusByDate[employee].push({ 
              isRemote, 
              id: recordId, 
              role: host === employee ? '主持人' : '陪餐人员',
              date: date
            });
          }
        }
      }

      // Rule 2: Split Travel Tracking
      if (enabled.splitTravel && expenseType.includes('差旅')) {
        const reason = String(record[m.travelReason] || '无事由');
        const key = `${employee}_${reason}`;
        if (!travelReasonMap[key]) travelReasonMap[key] = new Set();
        travelReasonMap[key].add(recordId);
      }

      // Rule 3: Invalid Guest
      if (enabled.invalidGuest) {
        const guestLower = String(record[m.guest] || '').toLowerCase();
        const invalidKeywords = currentConfig.invalidGuest.keywords;
        if (invalidKeywords.some(k => guestLower.includes(k.toLowerCase()))) {
          detectedRisks.push({
            recordId: recordId,
            ruleName: '招待对象异常',
            severity: 'medium',
            description: `招待对象 (${record[m.guest]}) 包含不符合商业招待必要性的身份（如司机）。`
          });
        }
      }

      // Rule 4: Internal Group Reception
      if (enabled.internalGroup) {
        const guestLower = String(record[m.guest] || '').toLowerCase();
        const descLower = String(record[m.description] || '').toLowerCase();
        const internalKeywords = currentConfig.internalGroup.keywords;
        if (internalKeywords.some(k => guestLower.includes(k.toLowerCase()) || descLower.includes(k.toLowerCase()))) {
          detectedRisks.push({
            recordId: recordId,
            ruleName: '集团内部接待',
            severity: 'high',
            description: `检测到集团内部接待记录（涉及关键词: ${internalKeywords.filter(k => guestLower.includes(k.toLowerCase()) || descLower.includes(k.toLowerCase())).join(', ')}）。`
          });
        }
      }
    });

    // Rule 1: Travel Location Mismatch (Final check)
    if (enabled.travelLocationMismatch) {
      Object.entries(receptionStatusByDate).forEach(([emp, receptions]) => {
        const travelDays = travelStatusByDate[emp] || {};
        receptions.forEach(rec => {
          const travelIsRemote = travelDays[rec.date];
          if (travelIsRemote !== undefined) {
            // If travel is remote ('是') but reception is local ('否')
            // Or travel is local ('否') but reception is remote ('是')
            if (travelIsRemote !== rec.isRemote) {
              detectedRisks.push({
                recordId: rec.id,
                ruleName: '出差与招待地点矛盾',
                severity: 'high',
                description: `该员工在 ${rec.date} 申请了${travelIsRemote === '是' ? '异地' : '本地'}出差，但同时作为${rec.isRemote === '是' ? '异地' : '本地'}招待的${rec.role}。`
              });
            }
          }
        });
      });
    }

    // Rule 2: Split Travel (Final check)
    if (enabled.splitTravel) {
      Object.entries(travelReasonMap).forEach(([key, orderIds]) => {
        if (orderIds.size > 1) {
          const [emp, reason] = key.split('_');
          orderIds.forEach(id => {
            detectedRisks.push({
              recordId: id,
              ruleName: '出差事由分拆报销',
              severity: 'medium',
              description: `员工 ${emp} 的出差事由 "${reason}" 存在多笔分拆报销记录（涉及单据: ${Array.from(orderIds).join(', ')}）。`
            });
          });
        }
      });
    }

    // Rule 2: Frequency Overload (Final check)
    if (enabled.frequency) {
      // Check daily limit
      filteredData.forEach(record => {
        const m = currentConfig.fieldMappings;
        const recordId = record[m.orderId] || record.id;
        const employee = String(record[m.employee] || '未知');
        const date = String(record[m.date] || '');
        const freqKey = `${employee}_${date}`;
        if (frequencyMap[freqKey] > currentConfig.frequency.limit) {
          detectedRisks.push({
            recordId: recordId,
            ruleName: '报销频次异常',
            severity: 'medium',
            description: `该员工在同一天 (${date}) 提交了 ${frequencyMap[freqKey]} 笔报销记录（超限 ${currentConfig.frequency.limit}）。`
          });
        }
      });

      // Check consecutive days
      const isWeekendOrHoliday = (ts: number) => {
        const d = new Date(ts);
        const day = d.getDay();
        if (day === 0 || day === 6) return true;
        const localDateStr = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
        return currentConfig.frequency.holidays.includes(localDateStr);
      };

      Object.entries(employeeDates).forEach(([emp, datesSet]) => {
        const dates = Array.from(datesSet)
          .map(d => new Date(d).getTime())
          .filter(t => !isNaN(t))
          .sort((a, b) => a - b);

        if (dates.length >= currentConfig.frequency.consecutiveDays) {
          let maxConsecutive = 1;
          let currentConsecutive = 1;
          let consecutiveSequence: number[] = [dates[0]];
          let maxSequence: number[] = [dates[0]];

          for (let i = 1; i < dates.length; i++) {
            const d1 = new Date(new Date(dates[i-1]).setHours(0,0,0,0)).getTime();
            const d2 = new Date(new Date(dates[i]).setHours(0,0,0,0)).getTime();
            const diffDays = Math.round((d2 - d1) / (1000 * 60 * 60 * 24));

            if (diffDays === 1) {
              currentConsecutive++;
              consecutiveSequence.push(dates[i]);
            } else if (diffDays > 1) {
              let allOffDays = true;
              for (let j = 1; j < diffDays; j++) {
                const midDay = d1 + j * 24 * 60 * 60 * 1000;
                if (!isWeekendOrHoliday(midDay)) {
                  allOffDays = false;
                  break;
                }
              }
              if (allOffDays) {
                currentConsecutive++;
                consecutiveSequence.push(dates[i]);
              } else {
                if (currentConsecutive > maxConsecutive) {
                  maxConsecutive = currentConsecutive;
                  maxSequence = [...consecutiveSequence];
                }
                currentConsecutive = 1;
                consecutiveSequence = [dates[i]];
              }
            }
          }
          if (currentConsecutive > maxConsecutive) {
            maxConsecutive = currentConsecutive;
            maxSequence = [...consecutiveSequence];
          }

          if (maxConsecutive >= currentConfig.frequency.consecutiveDays) {
            // Find records that fall into this sequence to flag them
            const startStr = new Date(maxSequence[0]).toLocaleDateString();
            const endStr = new Date(maxSequence[maxSequence.length - 1]).toLocaleDateString();
            
            filteredData.forEach(record => {
              const m = currentConfig.fieldMappings;
              const recordEmp = String(record[m.employee] || '未知');
              const recordDate = new Date(String(record[m.date] || '')).getTime();
              
              if (recordEmp === emp && maxSequence.includes(recordDate)) {
                detectedRisks.push({
                  recordId: record[m.orderId] || record.id,
                  ruleName: '连续报销异常',
                  severity: 'medium',
                  description: `该员工在 ${startStr} 至 ${endStr} 期间连续 ${maxConsecutive} 个工作日发生招待（超限 ${currentConfig.frequency.consecutiveDays} 天），已自动跳过周末及法定节假日。`
                });
              }
            });
          }
        }
      });
    }
    
    // Top N Employee Analysis
    if (enabled.topNEmployee) {
      const sortedEmployees = Object.entries(employeeTotals).sort((a, b) => b[1] - a[1]);
      const topN = sortedEmployees.slice(0, currentConfig.topN.n);
      const others = sortedEmployees.slice(currentConfig.topN.n);
      const avgOthers = others.length > 0 ? others.reduce((acc, curr) => acc + curr[1], 0) / others.length : 0;
      
      topN.forEach(([emp, total]) => {
        if (total > avgOthers * currentConfig.topN.thresholdMultiplier && avgOthers > 0) {
          // Find first record for this employee to attach risk
          const firstRec = filteredData.find(r => String(r[currentConfig.fieldMappings.employee]) === emp);
          if (firstRec) {
            detectedRisks.push({
              recordId: firstRec[currentConfig.fieldMappings.orderId] || firstRec.id,
              ruleName: '个人报销金额异常',
              severity: 'medium',
              description: `该员工累计报销金额 (${total.toFixed(2)}) 位居前 ${currentConfig.topN.n}，且超过其他员工平均水平的 ${currentConfig.topN.thresholdMultiplier} 倍。`
            });
          }
        }
      });
    }

    // Top N Department Analysis
    if (enabled.topNDepartment) {
      const sortedDepts = Object.entries(deptTotals).sort((a, b) => b[1] - a[1]);
      const topN = sortedDepts.slice(0, currentConfig.topN.n);
      const others = sortedDepts.slice(currentConfig.topN.n);
      const avgOthers = others.length > 0 ? others.reduce((acc, curr) => acc + curr[1], 0) / others.length : 0;
      
      topN.forEach(([dept, total]) => {
        if (total > avgOthers * currentConfig.topN.thresholdMultiplier && avgOthers > 0) {
          const firstRec = filteredData.find(r => String(r[currentConfig.fieldMappings.department]) === dept);
          if (firstRec) {
            detectedRisks.push({
              recordId: firstRec[currentConfig.fieldMappings.orderId] || firstRec.id,
              ruleName: '部门报销金额异常',
              severity: 'medium',
              description: `该部门 (${dept}) 累计报销金额 (${total.toFixed(2)}) 位居前 ${currentConfig.topN.n}，且超过其他部门平均水平的 ${currentConfig.topN.thresholdMultiplier} 倍。`,
              entityName: dept
            });
          }
        }
      });
    }
    
    // Travel Overlap Analysis
    if (enabled.travelOverlap) {
      Object.entries(travelIntervals).forEach(([emp, intervals]) => {
        for (let i = 0; i < intervals.length; i++) {
          for (let j = i + 1; j < intervals.length; j++) {
            const a = intervals[i];
            const b = intervals[j];
            // Check overlap: max(start) < min(end)
            if (Math.max(a.start, b.start) < Math.min(a.end, b.end)) {
              detectedRisks.push({
                recordId: a.id,
                ruleName: '差旅重复报销',
                severity: 'high',
                description: `该员工存在差旅时间重叠。记录 ${a.id} (${a.date}) 与 记录 ${b.id} (${b.date}) 的出差期间存在交叉。`
              });
            }
          }
        }
      });
    }

    setRisks(detectedRisks);
    setIsAuditOutdated(false);
    setIsProcessing(false);
    resolve();
      }, 50);
    });
  };

  const exportCombinedExcel = () => {
    if (records.length === 0 || risks.length === 0) return;

    // Create a map of risks by recordId for fast lookup
    const riskMap = new Map<string, AuditRisk[]>();
    for (const risk of risks) {
      if (!riskMap.has(risk.recordId)) {
        riskMap.set(risk.recordId, []);
      }
      riskMap.get(risk.recordId)!.push(risk);
    }

    // Filter records to only include those with risks
    const recordsWithRisks = records.filter(record => {
      const recordId = String(record[config.fieldMappings.orderId] || record.id);
      return riskMap.has(recordId);
    });

    const exportData = recordsWithRisks.map(record => {
      const recordId = String(record[config.fieldMappings.orderId] || record.id);
      const recordRisks = riskMap.get(recordId) || [];
      
      const exportRow = { ...record };
      
      return {
        ...exportRow,
        '命中的风险规则': recordRisks.map(r => r.ruleName).join('; ')
      };
    });

    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "审计分析结果");
    const fileName = `EagleEye_审计分析结果_${new Date().getTime()}.xlsx`;

    const isTauri =
      typeof window !== 'undefined' &&
      (('__TAURI_INTERNALS__' in window) || ('__TAURI__' in window));

    if (!isTauri) {
      XLSX.writeFile(wb, fileName);
      return;
    }

    (async () => {
      try {
        const [{save}, {writeFile}] = await Promise.all([
          import('@tauri-apps/plugin-dialog'),
          import('@tauri-apps/plugin-fs'),
        ]);

        const selectedPath = await save({
          defaultPath: fileName,
          filters: [{name: 'Excel', extensions: ['xlsx']}],
        });
        if (!selectedPath) return;

        const arrayBuffer = XLSX.write(wb, {
          bookType: 'xlsx',
          type: 'array',
        }) as ArrayBuffer;

        await writeFile(selectedPath, new Uint8Array(arrayBuffer));
      } catch (e) {
        console.error(e);
        alert('导出失败：请检查保存路径权限，或稍后重试。');
      }
    })();
  };

  const clearAll = () => {
    setRecords([]);
    setRisks([]);
    setFileName(null);
    setSearchTerm('');
    setAvailableExpenseTypes([]);
    setSelectedExpenseTypes([]);
  };

  const [searchTerm, setSearchTerm] = useState('');
  const [currentPage, setCurrentPage] = useState(1);
  const itemsPerPage = 50;

  const recordMap = useMemo(() => {
    const map = new Map();
    for (const r of records) {
      map.set(String(r[config.fieldMappings.orderId] || r.id), r);
    }
    return map;
  }, [records, config.fieldMappings.orderId]);

  const filteredRisks = useMemo(() => {
    const lowerSearchTerm = searchTerm.toLowerCase();
    
    return risks.filter(risk => {
      if (!lowerSearchTerm) return true;
      const record = recordMap.get(String(risk.recordId));
      const employee = record ? String(record[config.fieldMappings.employee] || '') : '';
      const searchStr = `${risk.recordId} ${employee} ${risk.ruleName} ${risk.description}`.toLowerCase();
      return searchStr.includes(lowerSearchTerm);
    });
  }, [risks, recordMap, searchTerm, config.fieldMappings.employee]);

  const riskDistribution = useMemo(() => {
    const dist = risks.reduce((acc, risk) => {
      acc[risk.ruleName] = (acc[risk.ruleName] || 0) + 1;
      return acc;
    }, {} as Record<string, number>);
    return Object.entries(dist).sort((a, b) => b[1] - a[1]);
  }, [risks]);

  const totalPages = Math.ceil(filteredRisks.length / itemsPerPage);
  const paginatedRisks = useMemo(() => {
    const start = (currentPage - 1) * itemsPerPage;
    return filteredRisks.slice(start, start + itemsPerPage);
  }, [filteredRisks, currentPage]);

  const updateConfig = (newConfig: Partial<AuditConfig>) => {
    const updated = { ...config, ...newConfig };
    setConfig(updated);
    setIsAuditOutdated(true);
  };

  const toggleExpenseType = (type: string) => {
    const newSelected = selectedExpenseTypes.includes(type)
      ? selectedExpenseTypes.filter(t => t !== type)
      : [...selectedExpenseTypes, type];
    setSelectedExpenseTypes(newSelected);
    setIsAuditOutdated(true);
  };

  const selectAllExpenseTypes = () => {
    setSelectedExpenseTypes(availableExpenseTypes);
    setIsAuditOutdated(true);
  };

  const deselectAllExpenseTypes = () => {
    setSelectedExpenseTypes([]);
    setIsAuditOutdated(true);
  };

  const handleExpenseTypeFieldChange = (e: React.ChangeEvent<HTMLSelectElement>) => {
    const newField = e.target.value;
    const newConfig = {
      ...config,
      fieldMappings: {
        ...config.fieldMappings,
        expenseType: newField
      }
    };
    setConfig(newConfig);
    
    if (records.length > 0) {
      const types = Array.from(new Set(records.map(r => String(r[newField] || '')).filter(Boolean)));
      setAvailableExpenseTypes(types);
      setSelectedExpenseTypes(types);
      setIsAuditOutdated(true);
    }
  };

  const loadSampleData = () => {
    const sampleRecords: ReimbursementRecord[] = [
      {
        id: 'BX2026001',
        机构: '总行',
        部门: '金融市场部',
        责任中心: '1001',
        报销人: '张三',
        费用所属员工: '张三',
        收款人: '某KTV',
        费用所属客户: '某供应商',
        '党组织/党务部门': '无',
        报销单编号: 'BX2026001',
        费用类型: '招待费',
        费用科目: '业务招待费',
        '费用描述（招待原因）': '与技术供应商在KTV进行业务洽谈',
        币种: 'CNY',
        原币金额: 1200,
        折人民币金额: 1200,
        反冲金额: 0,
        税额: 0,
        状态: '已审批',
        发生日期: '2026-03-20',
        提交日期: '2026-03-21',
        记账日期: '2026-03-22',
        当前审批人: '系统',
        已审批人: '主管A',
        审批日期: '2026-03-22',
        审批信息: '通过',
        附件张数: 3,
        费用终止日期: '2026-03-20',
        出差天数: 0,
        旅差费开始结束日期: '',
        应酬招待主持人: '张三',
        主持人人数: 1,
        行内陪餐人员: '李四,王五',
        行内陪餐人员人数: 4,
        工作人员: '赵六',
        工作人员人数: 1,
        招待对象: '技术供应商A',
        招待对象人数: 2,
        招待人均金额: 600,
        工作人员人均金额: 0,
        赠品名称: '无',
        接待类型: '商务接待',
        是否领用行内酒水: '否',
        酒水金额: 1200,
        是否异地: '否',
        出差地点: '',
        契约号: '',
        关联发票: 'INV001',
        HR系统审批单号: 'HR001',
        特殊情况说明: '无'
      },
      {
        id: 'BX2026002',
        机构: '总行',
        部门: '投资银行部',
        责任中心: '1002',
        报销人: '李四',
        费用所属员工: '李四',
        收款人: '某餐厅',
        费用所属客户: '某投资机构',
        '党组织/党务部门': '无',
        报销单编号: 'BX2026002',
        费用类型: '招待费',
        费用科目: '业务招待费',
        '费用描述（招待原因）': '商务午餐',
        币种: 'CNY',
        原币金额: 300,
        折人民币金额: 300,
        反冲金额: 0,
        税额: 0,
        状态: '已审批',
        发生日期: '2026-03-21',
        提交日期: '2026-03-21',
        记账日期: '2026-03-22',
        当前审批人: '系统',
        已审批人: '主管B',
        审批日期: '2026-03-22',
        审批信息: '通过',
        附件张数: 1,
        费用终止日期: '2026-03-21',
        出差天数: 0,
        旅差费开始结束日期: '',
        应酬招待主持人: '司机王五',
        主持人人数: 1,
        行内陪餐人员: '李四',
        行内陪餐人员人数: 1,
        工作人员: '无',
        工作人员人数: 0,
        招待对象: '某投资机构代表',
        招待对象人数: 2,
        招待人均金额: 150,
        工作人员人均金额: 0,
        赠品名称: '无',
        接待类型: '商务接待',
        是否领用行内酒水: '是',
        酒水金额: 0,
        是否异地: '否',
        出差地点: '',
        契约号: '',
        关联发票: 'INV002',
        HR系统审批单号: 'HR002',
        特殊情况说明: '无'
      },
      {
        id: 'BX2026003',
        机构: '总行',
        部门: '零售银行部',
        责任中心: '1003',
        报销人: '王五',
        费用所属员工: '王五',
        收款人: '某餐厅',
        费用所属客户: '个人客户',
        '党组织/党务部门': '无',
        报销单编号: 'BX2026003',
        费用类型: '招待费',
        费用科目: '业务招待费',
        '费用描述（招待原因）': '商务午餐',
        币种: 'CNY',
        原币金额: 1600,
        折人民币金额: 1600,
        反冲金额: 0,
        税额: 0,
        状态: '已审批',
        发生日期: '2026-03-21',
        提交日期: '2026-03-21',
        记账日期: '2026-03-22',
        当前审批人: '系统',
        已审批人: '主管C',
        审批日期: '2026-03-22',
        审批信息: '通过',
        附件张数: 2,
        费用终止日期: '2026-03-21',
        出差天数: 0,
        旅差费开始结束日期: '',
        应酬招待主持人: '王五',
        主持人人数: 1,
        行内陪餐人员: '赵六',
        行内陪餐人员人数: 1,
        工作人员: '无',
        工作人员人数: 0,
        招待对象: '某重要客户',
        招待对象人数: 1,
        招待人均金额: 800,
        工作人员人均金额: 0,
        赠品名称: '无',
        接待类型: '商务接待',
        是否领用行内酒水: '否',
        酒水金额: 0,
        是否异地: '否',
        出差地点: '',
        契约号: '',
        关联发票: 'INV003',
        HR系统审批单号: 'HR003',
        特殊情况说明: '无'
      },
      {
        id: 'BX2026004',
        机构: '总行',
        部门: '零售银行部',
        责任中心: '1003',
        报销人: '王五',
        费用所属员工: '王五',
        收款人: '某餐厅',
        费用所属客户: '个人客户',
        '党组织/党务部门': '无',
        报销单编号: 'BX2026004',
        费用类型: '招待费',
        费用科目: '业务招待费',
        '费用描述（招待原因）': '商务晚餐',
        币种: 'CNY',
        原币金额: 1200,
        折人民币金额: 1200,
        反冲金额: 0,
        税额: 0,
        状态: '已审批',
        发生日期: '2026-03-21',
        提交日期: '2026-03-21',
        记账日期: '2026-03-22',
        当前审批人: '系统',
        已审批人: '主管C',
        审批日期: '2026-03-22',
        审批信息: '通过',
        附件张数: 2,
        费用终止日期: '2026-03-21',
        出差天数: 0,
        旅差费开始结束日期: '',
        应酬招待主持人: '王五',
        主持人人数: 1,
        行内陪餐人员: '赵六',
        行内陪餐人员人数: 1,
        工作人员: '无',
        工作人员人数: 0,
        招待对象: '某重要客户',
        招待对象人数: 1,
        招待人均金额: 600,
        工作人员人均金额: 0,
        赠品名称: '无',
        接待类型: '商务接待',
        是否领用行内酒水: '否',
        酒水金额: 0,
        是否异地: '否',
        出差地点: '',
        契约号: '',
        关联发票: 'INV004',
        HR系统审批单号: 'HR004',
        特殊情况说明: '无'
      },
      {
        id: 'BX2026005',
        机构: '总行',
        部门: '零售银行部',
        责任中心: '1003',
        报销人: '王五',
        费用所属员工: '王五',
        收款人: '某餐厅',
        费用所属客户: '个人客户',
        '党组织/党务部门': '无',
        报销单编号: 'BX2026005',
        费用类型: '招待费',
        费用科目: '业务招待费',
        '费用描述（招待原因）': '商务晚餐',
        币种: 'CNY',
        原币金额: 1000,
        折人民币金额: 1000,
        反冲金额: 0,
        税额: 0,
        状态: '已审批',
        发生日期: '2026-03-21',
        提交日期: '2026-03-21',
        记账日期: '2026-03-22',
        当前审批人: '系统',
        已审批人: '主管C',
        审批日期: '2026-03-22',
        审批信息: '通过',
        附件张数: 2,
        费用终止日期: '2026-03-21',
        出差天数: 0,
        旅差费开始结束日期: '',
        应酬招待主持人: '王五',
        主持人人数: 1,
        行内陪餐人员: '赵六',
        行内陪餐人员人数: 1,
        工作人员: '无',
        工作人员人数: 0,
        招待对象: '某重要客户',
        招待对象人数: 1,
        招待人均金额: 500,
        工作人员人均金额: 0,
        赠品名称: '无',
        接待类型: '商务接待',
        是否领用行内酒水: '否',
        酒水金额: 0,
        是否异地: '否',
        出差地点: '',
        契约号: '',
        关联发票: 'INV005',
        HR系统审批单号: 'HR005',
        特殊情况说明: '无'
      },
      {
        id: 'BX2026006',
        机构: '总行',
        部门: '公司金融部',
        责任中心: '1004',
        报销人: '张三',
        费用所属员工: '张三',
        收款人: '某酒店',
        费用所属客户: '某国企总部',
        '党组织/党务部门': '无',
        报销单编号: 'BX2026006',
        费用类型: '招待费',
        费用科目: '业务招待费',
        '费用描述（招待原因）': '国企总部考察接待',
        币种: 'CNY',
        原币金额: 2000,
        折人民币金额: 2000,
        反冲金额: 0,
        税额: 0,
        状态: '已审批',
        发生日期: '2026-03-05',
        提交日期: '2026-03-06',
        记账日期: '2026-03-07',
        当前审批人: '系统',
        已审批人: '主管D',
        审批日期: '2026-03-07',
        审批信息: '通过',
        附件张数: 5,
        费用终止日期: '2026-03-05',
        出差天数: 0,
        旅差费开始结束日期: '',
        应酬招待主持人: '张三',
        主持人人数: 1,
        行内陪餐人员: '李四',
        行内陪餐人员人数: 1,
        工作人员: '无',
        工作人员人数: 0,
        招待对象: '国企总部领导',
        招待对象人数: 3,
        招待人均金额: 666,
        工作人员人均金额: 0,
        赠品名称: '无',
        接待类型: '商务接待',
        是否领用行内酒水: '否',
        酒水金额: 0,
        是否异地: '否',
        出差地点: '',
        契约号: '',
        关联发票: 'INV006',
        HR系统审批单号: 'HR006',
        特殊情况说明: '休假期间报销',
        出差事由: '',
        陪餐人员: '李四'
      },
      {
        id: 'BX2026007',
        机构: '总行',
        部门: '差旅部',
        责任中心: '1005',
        报销人: '王五',
        费用所属员工: '王五',
        收款人: '某酒店',
        费用所属客户: '无',
        '党组织/党务部门': '无',
        报销单编号: 'BX2026007',
        费用类型: '差旅费',
        费用科目: '交通费',
        '费用描述（招待原因）': '出差考察',
        币种: 'CNY',
        原币金额: 500,
        折人民币金额: 500,
        反冲金额: 0,
        税额: 0,
        状态: '已审批',
        发生日期: '2026-03-10',
        提交日期: '2026-03-11',
        记账日期: '2026-03-12',
        当前审批人: '系统',
        已审批人: '主管E',
        审批日期: '2026-03-12',
        审批信息: '通过',
        附件张数: 1,
        费用终止日期: '2026-03-15',
        出差天数: 5,
        旅差费开始结束日期: '2026-03-10至2026-03-15',
        应酬招待主持人: '',
        主持人人数: 0,
        行内陪餐人员: '',
        行内陪餐人员人数: 0,
        工作人员: '',
        工作人员人数: 0,
        招待对象: '',
        招待对象人数: 0,
        招待人均金额: 0,
        工作人员人均金额: 0,
        赠品名称: '',
        接待类型: '',
        是否领用行内酒水: '否',
        酒水金额: 0,
        是否异地: '是',
        出差地点: '上海',
        契约号: '',
        关联发票: 'INV007',
        HR系统审批单号: 'HR007',
        特殊情况说明: '重复报销测试'
      },
      {
        id: 'BX2026008',
        机构: '总行',
        部门: '差旅部',
        责任中心: '1005',
        报销人: '王五',
        费用所属员工: '王五',
        收款人: '某酒店',
        费用所属客户: '无',
        '党组织/党务部门': '无',
        报销单编号: 'BX2026008',
        费用类型: '差旅费',
        费用科目: '交通费',
        '费用描述（招待原因）': '出差考察2',
        币种: 'CNY',
        原币金额: 600,
        折人民币金额: 600,
        反冲金额: 0,
        税额: 0,
        状态: '已审批',
        发生日期: '2026-03-12',
        提交日期: '2026-03-13',
        记账日期: '2026-03-14',
        当前审批人: '系统',
        已审批人: '主管E',
        审批日期: '2026-03-14',
        审批信息: '通过',
        附件张数: 1,
        费用终止日期: '2026-03-17',
        出差天数: 5,
        旅差费开始结束日期: '2026-03-12至2026-03-17',
        应酬招待主持人: '',
        主持人人数: 0,
        行内陪餐人员: '',
        行内陪餐人员人数: 0,
        工作人员: '',
        工作人员人数: 0,
        招待对象: '',
        招待对象人数: 0,
        招待人均金额: 0,
        工作人员人均金额: 0,
        赠品名称: '',
        接待类型: '',
        是否领用行内酒水: '否',
        酒水金额: 0,
        是否异地: '是',
        出差地点: '北京',
        契约号: '',
        关联发票: 'INV008',
        HR系统审批单号: 'HR008',
        特殊情况说明: '重复报销测试',
        出差事由: '项目调研',
        陪餐人员: ''
      },
      {
        id: 'BX2026009',
        机构: '总行',
        部门: '业务部',
        责任中心: '1006',
        报销人: '赵六',
        费用所属员工: '赵六',
        收款人: '某餐厅',
        费用所属客户: '客户H',
        '党组织/党务部门': '否',
        报销单编号: 'BX2026009',
        费用类型: '业务招待费',
        费用科目: '餐费',
        '费用描述（招待原因）': '业务洽谈',
        币种: 'CNY',
        原币金额: 800,
        折人民币金额: 800,
        反冲金额: 0,
        税额: 0,
        状态: '已审批',
        发生日期: '2026-03-20',
        提交日期: '2026-03-21',
        记账日期: '2026-03-22',
        当前审批人: '系统',
        已审批人: '经理C',
        审批日期: '2026-03-22',
        审批信息: '通过',
        附件张数: 1,
        费用终止日期: '',
        出差天数: 0,
        旅差费开始结束日期: '',
        应酬招待主持人: '赵六',
        主持人人数: 1,
        行内陪餐人员: '',
        行内陪餐人员人数: 0,
        工作人员: '',
        工作人员人数: 0,
        招待对象: '司机师傅',
        招待对象人数: 1,
        招待人均金额: 800,
        工作人员人均金额: 0,
        赠品名称: '',
        接待类型: '商务接待',
        是否领用行内酒水: '否',
        酒水金额: 0,
        是否异地: '否',
        出差地点: '',
        契约号: '',
        关联发票: 'INV009',
        HR系统审批单号: 'HR009',
        特殊情况说明: '',
        出差事由: '',
        陪餐人员: ''
      },
      {
        id: 'BX2026010',
        机构: '总行',
        部门: '业务部',
        责任中心: '1006',
        报销人: '钱七',
        费用所属员工: '钱七',
        收款人: '某餐厅',
        费用所属客户: '客户I',
        '党组织/党务部门': '否',
        报销单编号: 'BX2026010',
        费用类型: '业务招待费',
        费用科目: '餐费',
        '费用描述（招待原因）': 'xib内部交流',
        币种: 'CNY',
        原币金额: 1200,
        折人民币金额: 1200,
        反冲金额: 0,
        税额: 0,
        状态: '已审批',
        发生日期: '2026-03-25',
        提交日期: '2026-03-26',
        记账日期: '2026-03-27',
        当前审批人: '系统',
        已审批人: '经理C',
        审批日期: '2026-03-27',
        审批信息: '通过',
        附件张数: 1,
        费用终止日期: '',
        出差天数: 0,
        旅差费开始结束日期: '',
        应酬招待主持人: '钱七',
        主持人人数: 1,
        行内陪餐人员: '',
        行内陪餐人员人数: 0,
        工作人员: '',
        工作人员人数: 0,
        招待对象: 'xib同事',
        招待对象人数: 2,
        招待人均金额: 400,
        工作人员人均金额: 0,
        赠品名称: '',
        接待类型: '商务接待',
        是否领用行内酒水: '否',
        酒水金额: 0,
        是否异地: '否',
        出差地点: '',
        契约号: '',
        关联发票: 'INV010',
        HR系统审批单号: 'HR010',
        特殊情况说明: '',
        出差事由: '',
        陪餐人员: ''
      }
    ];
    
    setFileName('示例数据.xlsx');
    setHeaders(DEFAULT_HEADERS);
    setRecords(sampleRecords);
    
    const types = Array.from(new Set(sampleRecords.map(r => r[config.fieldMappings.expenseType]).filter(Boolean)));
    setAvailableExpenseTypes(types);
    setSelectedExpenseTypes(types);
    
    runAudit(sampleRecords, config, types);
  };

  return (
    <div className="min-h-screen bg-[#E4E3E0] text-[#141414] font-sans selection:bg-[#141414] selection:text-[#E4E3E0]">
      {/* Toast Notification */}
      {toastMessage && (
        <div className="fixed top-4 left-1/2 -translate-x-1/2 z-50">
          <motion.div 
            initial={{ opacity: 0, y: -20 }}
            animate={{ opacity: 1, y: 0 }}
            exit={{ opacity: 0, y: -20 }}
            className={cn(
              "px-6 py-3 shadow-xl border flex items-center gap-3 font-bold text-sm",
              toastMessage.type === 'success' ? "bg-green-50 text-green-800 border-green-200" : "bg-red-50 text-red-800 border-red-200"
            )}
          >
            {toastMessage.type === 'success' ? <CheckCircle2 size={18} /> : <AlertTriangle size={18} />}
            {toastMessage.message}
          </motion.div>
        </div>
      )}

      {/* Header */}
      <header className="border-b border-[#141414] p-6 flex justify-between items-center bg-[#E4E3E0] sticky top-0 z-10">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 bg-[#141414] flex items-center justify-center rounded-sm">
            <Shield className="text-[#E4E3E0] w-6 h-6" />
          </div>
          <div>
            <h1 className="text-xl font-bold tracking-tighter uppercase">EagleEye 审计工具</h1>
            <p className="text-[10px] font-mono opacity-50 uppercase tracking-widest">Win10 x64 离线审计核心 v1.0</p>
          </div>
        </div>
        
        <div className="flex gap-4">
          <button 
            onClick={loadSampleData}
            className="flex items-center gap-2 px-4 py-2 border border-[#141414] border-dashed hover:bg-[#141414]/5 transition-all text-xs font-bold uppercase tracking-tighter"
          >
            加载示例数据
          </button>
          <button 
            onClick={() => fileInputRef.current?.click()}
            className="flex items-center gap-2 px-4 py-2 border border-[#141414] hover:bg-[#141414] hover:text-[#E4E3E0] transition-all text-xs font-bold uppercase tracking-tighter"
          >
            <FileUp size={16} />
            导入 Excel
          </button>
          <input 
            type="file" 
            ref={fileInputRef} 
            onChange={handleFileUpload} 
            accept=".xlsx, .xls" 
            className="hidden" 
          />
          
          <button 
            onClick={exportCombinedExcel}
            disabled={records.length === 0 || risks.length === 0}
            className="flex items-center gap-2 px-4 py-2 bg-[#141414] text-[#E4E3E0] hover:opacity-90 disabled:opacity-30 transition-all text-xs font-bold uppercase tracking-tighter"
          >
            <Table size={16} />
            导出分析结果
          </button>

          {records.length > 0 && (
            <button 
              onClick={clearAll}
              className="flex items-center gap-2 px-4 py-2 border border-red-600 text-red-600 hover:bg-red-600 hover:text-white transition-all text-xs font-bold uppercase tracking-tighter"
            >
              <Trash2 size={16} />
              清空
            </button>
          )}
        </div>
      </header>

      <main className="p-6 max-w-7xl mx-auto">
        {/* Risk Rules Configuration */}
        <div className="mb-8 border border-[#141414] bg-white">
          <button 
            onClick={() => setShowConfig(!showConfig)}
            className="w-full p-4 flex justify-between items-center bg-[#f5f5f5] hover:bg-[#eee] transition-colors"
          >
            <div className="flex items-center gap-2">
              <Filter size={16} />
              <h2 className="font-serif italic text-sm uppercase tracking-wider">风险规则自定义配置</h2>
            </div>
            <span className="text-[10px] font-mono opacity-50 uppercase tracking-widest">
              {showConfig ? '收起配置 [-]' : '展开配置 [+]'}
            </span>
          </button>
          
          <AnimatePresence>
            {showConfig && (
              <motion.div 
                initial={{ height: 0, opacity: 0 }}
                animate={{ height: 'auto', opacity: 1 }}
                exit={{ height: 0, opacity: 0 }}
                className="overflow-hidden"
              >
                <div className="p-6 border-t border-[#141414]">
                  <div className="flex justify-between items-center mb-6">
                    <h3 className="text-xs font-bold uppercase tracking-tighter">规则启用与参数配置</h3>
                    <div className="flex gap-2">
                      <button 
                        onClick={() => {
                          const allEnabled = Object.keys(config.enabledRules).reduce((acc, key) => ({ ...acc, [key]: true }), {});
                          updateConfig({ enabledRules: allEnabled as any });
                        }}
                        className="text-[10px] font-mono uppercase bg-[#141414] text-[#E4E3E0] px-2 py-1 hover:opacity-80 transition-opacity"
                      >
                        全选
                      </button>
                      <button 
                        onClick={() => {
                          const allDisabled = Object.keys(config.enabledRules).reduce((acc, key) => ({ ...acc, [key]: false }), {});
                          updateConfig({ enabledRules: allDisabled as any });
                        }}
                        className="text-[10px] font-mono uppercase border border-[#141414] px-2 py-1 hover:bg-[#141414]/5 transition-colors"
                      >
                        全不选
                      </button>
                    </div>
                  </div>

                  <div className="space-y-6">
                    {RULE_DEFINITIONS.map((rule) => (
                      <div key={rule.key} className={cn(
                        "p-4 border border-[#141414]/10 transition-all",
                        config.enabledRules[rule.key as keyof typeof config.enabledRules] ? "bg-white border-[#141414]" : "bg-gray-50 opacity-60"
                      )}>
                        <div className="flex items-start gap-4">
                          <div className="pt-1">
                            <input 
                              type="checkbox" 
                              checked={config.enabledRules[rule.key as keyof typeof config.enabledRules]}
                              onChange={(e) => updateConfig({ enabledRules: { ...config.enabledRules, [rule.key]: e.target.checked } })}
                              className="w-5 h-5 accent-[#141414] cursor-pointer"
                              id={`rule-${rule.key}`}
                            />
                          </div>
                          <div className="flex-1">
                            <div className="flex justify-between items-start mb-2">
                              <div>
                                <label htmlFor={`rule-${rule.key}`} className="text-sm font-bold uppercase tracking-tight cursor-pointer block mb-1">
                                  {rule.name}
                                </label>
                                <p className="text-[11px] leading-relaxed opacity-70 mb-2">
                                  {rule.description}
                                </p>
                                <div className="bg-[#141414]/5 p-2 rounded-sm mb-3">
                                  <p className="text-[10px] font-mono italic opacity-60">
                                    {rule.example}
                                  </p>
                                  <p className="text-[9px] font-mono mt-1 text-[#141414]/40">
                                    默认使用字段: {rule.defaultFields.map(f => `[${f}]`).join('、')}
                                  </p>
                                </div>
                              </div>
                              <span className="text-[9px] font-mono opacity-30 uppercase">Rule ID: {rule.key}</span>
                            </div>

                            {/* Rule Specific Parameters */}
                            {config.enabledRules[rule.key as keyof typeof config.enabledRules] && rule.params.length > 0 && (
                              <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mt-4 pt-4 border-t border-[#141414]/5">
                                {rule.key === 'ghostHost' && (
                                  <>
                                    <div className="space-y-1">
                                      <label className="text-[10px] font-mono uppercase opacity-60">主持人关键词</label>
                                      <input 
                                        type="text"
                                        value={config.ghostHost.hosts.join(', ')}
                                        onChange={(e) => updateConfig({ ghostHost: { ...config.ghostHost, hosts: e.target.value.split(',').map(s => s.trim()) } })}
                                        className="w-full p-2 border border-[#141414] text-[11px] font-mono focus:ring-1 focus:ring-[#141414]"
                                      />
                                    </div>
                                    <div className="space-y-1">
                                      <label className="text-[10px] font-mono uppercase opacity-60">敏感对象关键词</label>
                                      <input 
                                        type="text"
                                        value={config.ghostHost.targets.join(', ')}
                                        onChange={(e) => updateConfig({ ghostHost: { ...config.ghostHost, targets: e.target.value.split(',').map(s => s.trim()) } })}
                                        className="w-full p-2 border border-[#141414] text-[11px] font-mono focus:ring-1 focus:ring-[#141414]"
                                      />
                                    </div>
                                  </>
                                )}
                                {rule.key === 'frequency' && (
                                  <>
                                    <div className="space-y-1">
                                      <label className="text-[10px] font-mono uppercase opacity-60">每日报销单上限 (笔)</label>
                                      <input 
                                        type="number"
                                        value={config.frequency.limit}
                                        onChange={(e) => updateConfig({ frequency: { ...config.frequency, limit: parseInt(e.target.value) || 0 } })}
                                        className="w-full p-2 border border-[#141414] text-[11px] font-mono focus:ring-1 focus:ring-[#141414]"
                                      />
                                    </div>
                                    <div className="space-y-1">
                                      <label className="text-[10px] font-mono uppercase opacity-60">连续报销天数上限</label>
                                      <input 
                                        type="number"
                                        value={config.frequency.consecutiveDays}
                                        onChange={(e) => updateConfig({ frequency: { ...config.frequency, consecutiveDays: parseInt(e.target.value) || 0 } })}
                                        className="w-full p-2 border border-[#141414] text-[11px] font-mono focus:ring-1 focus:ring-[#141414]"
                                      />
                                    </div>
                                    <div className="space-y-1 md:col-span-2">
                                      <label className="text-[10px] font-mono uppercase opacity-60">法定节假日 (每行一个日期，格式: YYYY-MM-DD)</label>
                                      <textarea 
                                        rows={2}
                                        placeholder="2026-05-01&#10;2026-10-01"
                                        value={config.frequency.holidays.join('\n')}
                                        onChange={(e) => updateConfig({ frequency: { ...config.frequency, holidays: e.target.value.split('\n').map(s => s.trim()).filter(Boolean) } })}
                                        className="w-full p-2 border border-[#141414] text-[11px] font-mono focus:ring-1 focus:ring-[#141414]"
                                      />
                                    </div>
                                  </>
                                )}
                                {rule.key === 'venues' && (
                                  <div className="space-y-1 md:col-span-2">
                                    <label className="text-[10px] font-mono uppercase opacity-60">违规关键词 (逗号分隔)</label>
                                    <textarea 
                                      rows={2}
                                      value={config.venues.keywords.join(', ')}
                                      onChange={(e) => updateConfig({ venues: { keywords: e.target.value.split(',').map(s => s.trim()) } })}
                                      className="w-full p-2 border border-[#141414] text-[11px] font-mono focus:ring-1 focus:ring-[#141414]"
                                    />
                                  </div>
                                )}
                                {rule.key === 'perCapita' && (
                                  <>
                                    <div className="space-y-1">
                                      <label className="text-[10px] font-mono uppercase opacity-60">人均限额 (元)</label>
                                      <input 
                                        type="number"
                                        value={config.perCapita.limit}
                                        onChange={(e) => updateConfig({ perCapita: { ...config.perCapita, limit: parseFloat(e.target.value) || 0 } })}
                                        className="w-full p-2 border border-[#141414] text-[11px] font-mono focus:ring-1 focus:ring-[#141414]"
                                      />
                                    </div>
                                    <div className="space-y-1">
                                      <label className="text-[10px] font-mono uppercase opacity-60">陪餐比例 (陪:客)</label>
                                      <input 
                                        type="number"
                                        step="0.1"
                                        value={config.perCapita.ratio}
                                        onChange={(e) => updateConfig({ perCapita: { ...config.perCapita, ratio: parseFloat(e.target.value) || 0 } })}
                                        className="w-full p-2 border border-[#141414] text-[11px] font-mono focus:ring-1 focus:ring-[#141414]"
                                      />
                                    </div>
                                  </>
                                )}
                                {rule.key === 'drinking' && (
                                  <>
                                    <div className="space-y-1">
                                      <label className="text-[10px] font-mono uppercase opacity-60">"未领用"关键词</label>
                                      <input 
                                        type="text"
                                        value={config.drinking.noKeywords.join(', ')}
                                        onChange={(e) => updateConfig({ drinking: { ...config.drinking, noKeywords: e.target.value.split(',').map(s => s.trim()) } })}
                                        className="w-full p-2 border border-[#141414] text-[11px] font-mono focus:ring-1 focus:ring-[#141414]"
                                      />
                                    </div>
                                    <div className="space-y-1">
                                      <label className="text-[10px] font-mono uppercase opacity-60">酒水识别关键词</label>
                                      <input 
                                        type="text"
                                        value={config.drinking.alcoholKeywords.join(', ')}
                                        onChange={(e) => updateConfig({ drinking: { ...config.drinking, alcoholKeywords: e.target.value.split(',').map(s => s.trim()) } })}
                                        className="w-full p-2 border border-[#141414] text-[11px] font-mono focus:ring-1 focus:ring-[#141414]"
                                      />
                                    </div>
                                  </>
                                )}
                                {rule.key === 'soe' && (
                                  <div className="space-y-1 md:col-span-2">
                                    <label className="text-[10px] font-mono uppercase opacity-60">国企/总部关键词</label>
                                    <textarea 
                                      rows={2}
                                      value={config.soe.keywords.join(', ')}
                                      onChange={(e) => updateConfig({ soe: { keywords: e.target.value.split(',').map(s => s.trim()) } })}
                                      className="w-full p-2 border border-[#141414] text-[11px] font-mono focus:ring-1 focus:ring-[#141414]"
                                    />
                                  </div>
                                )}
                                {rule.key === 'leave' && (
                                  <div className="space-y-1 md:col-span-2">
                                    <div className="flex justify-between items-center">
                                      <label className="text-[10px] font-mono uppercase opacity-60">异常期间列表 (格式: 姓名:开始:结束:类型)</label>
                                      <div>
                                        <input 
                                          type="file" 
                                          ref={leaveFileInputRef} 
                                          onChange={handleLeaveFileUpload} 
                                          accept=".xlsx, .xls" 
                                          className="hidden" 
                                        />
                                        <button 
                                          onClick={() => leaveFileInputRef.current?.click()}
                                          className="text-[10px] font-mono uppercase border border-[#141414] px-2 py-1 hover:bg-[#141414]/5 transition-colors flex items-center gap-1"
                                        >
                                          <FileUp size={10} />
                                          导入休假明细
                                        </button>
                                      </div>
                                    </div>
                                    {config.leave.periods.length > 100 ? (
                                      <div className="w-full p-2 border border-[#141414]/20 text-[11px] font-mono bg-gray-50 flex justify-between items-center">
                                        <span>已导入 {config.leave.periods.length} 条记录。由于数据量较大，已禁用手动编辑。请重新导入 Excel 文件以更新。</span>
                                        <button 
                                          onClick={() => updateConfig({ leave: { periods: [] } })}
                                          className="text-red-600 hover:underline"
                                        >
                                          清空记录
                                        </button>
                                      </div>
                                    ) : (
                                      <textarea 
                                        rows={3}
                                        placeholder="张三:2026-03-01:2026-03-10:年假"
                                        value={config.leave.periods.map(p => `${p.employee}:${p.start}:${p.end}:${p.type}`).join('\n')}
                                        onChange={(e) => {
                                          const lines = e.target.value.split('\n').filter(Boolean);
                                          const periods = lines.map(l => {
                                            const parts = l.split(':');
                                            if (parts.length === 4) {
                                              const [employee, start, end, type] = parts;
                                              return { employee, start, end, type };
                                            }
                                            return null;
                                          }).filter(Boolean) as LeavePeriod[];
                                          updateConfig({ leave: { periods } });
                                        }}
                                        className="w-full p-2 border border-[#141414] text-[11px] font-mono focus:ring-1 focus:ring-[#141414]"
                                      />
                                    )}
                                  </div>
                                )}
                                {rule.key === 'invalidGuest' && (
                                  <div className="space-y-1 md:col-span-2">
                                    <label className="text-[10px] font-mono uppercase opacity-60">异常身份关键词 (逗号分隔)</label>
                                    <textarea 
                                      rows={2}
                                      value={config.invalidGuest.keywords.join(', ')}
                                      onChange={(e) => updateConfig({ invalidGuest: { keywords: e.target.value.split(',').map(s => s.trim()) } })}
                                      className="w-full p-2 border border-[#141414] text-[11px] font-mono focus:ring-1 focus:ring-[#141414]"
                                    />
                                  </div>
                                )}
                                {rule.key === 'internalGroup' && (
                                  <div className="space-y-1 md:col-span-2">
                                    <label className="text-[10px] font-mono uppercase opacity-60">内部接待关键词 (逗号分隔)</label>
                                    <textarea 
                                      rows={2}
                                      value={config.internalGroup.keywords.join(', ')}
                                      onChange={(e) => updateConfig({ internalGroup: { keywords: e.target.value.split(',').map(s => s.trim()) } })}
                                      className="w-full p-2 border border-[#141414] text-[11px] font-mono focus:ring-1 focus:ring-[#141414]"
                                    />
                                  </div>
                                )}
                                {(rule.key === 'topNEmployee' || rule.key === 'topNDepartment') && (
                                  <>
                                    <div className="space-y-1">
                                      <label className="text-[10px] font-mono uppercase opacity-60">Top N 数量</label>
                                      <input 
                                        type="number"
                                        value={config.topN.n}
                                        onChange={(e) => updateConfig({ topN: { ...config.topN, n: parseInt(e.target.value) || 0 } })}
                                        className="w-full p-2 border border-[#141414] text-[11px] font-mono focus:ring-1 focus:ring-[#141414]"
                                      />
                                    </div>
                                    <div className="space-y-1">
                                      <label className="text-[10px] font-mono uppercase opacity-60">偏离倍数阈值</label>
                                      <input 
                                        type="number"
                                        step="0.5"
                                        value={config.topN.thresholdMultiplier}
                                        onChange={(e) => updateConfig({ topN: { ...config.topN, thresholdMultiplier: parseFloat(e.target.value) || 0 } })}
                                        className="w-full p-2 border border-[#141414] text-[11px] font-mono focus:ring-1 focus:ring-[#141414]"
                                      />
                                    </div>
                                  </>
                                )}
                              </div>
                            )}
                          </div>
                        </div>
                      </div>
                    ))}
                  </div>

                  {/* Field Mapping Config */}
                  <div className="lg:col-span-3 pt-6 border-t border-[#141414]/10">
                    <div className="flex justify-between items-center mb-4">
                      <h3 className="text-xs font-bold uppercase tracking-tighter">10. 数据字段映射配置</h3>
                      <span className="text-[9px] font-mono opacity-40 italic">将规则逻辑关联至 Excel 表头字段</span>
                    </div>
                    <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-5 gap-4">
                      {[
                        { key: 'employee', label: '费用所属员工' },
                        { key: 'department', label: '部门' },
                        { key: 'amount', label: '折人民币金额' },
                        { key: 'date', label: '发生日期' },
                        { key: 'host', label: '应酬招待主持人' },
                        { key: 'guest', label: '招待对象' },
                        { key: 'description', label: '费用描述' },
                        { key: 'alcoholStatus', label: '酒水领用状态' },
                        { key: 'alcoholAmount', label: '酒水报销金额' },
                        { key: 'perCapita', label: '招待人均金额' },
                        { key: 'guestCount', label: '招待对象人数' },
                        { key: 'staffCount', label: '陪餐人员人数' },
                        { key: 'receptionType', label: '接待类型' },
                        { key: 'travelDays', label: '出差天数' },
                        { key: 'orderId', label: '报销单编号' },
                        { key: 'isRemote', label: '是否异地' },
                        { key: 'expenseType', label: '费用类型' },
                        { key: 'travelReason', label: '出差事由' },
                        { key: 'staffNames', label: '陪餐人员' },
                      ].map((mapping) => (
                        <div key={mapping.key} className="space-y-1">
                          <label className="text-[9px] font-mono uppercase opacity-50 block">{mapping.label}</label>
                          <select 
                            value={config.fieldMappings[mapping.key as keyof typeof config.fieldMappings]}
                            onChange={(e) => updateConfig({ 
                              fieldMappings: { 
                                ...config.fieldMappings, 
                                [mapping.key]: e.target.value 
                              } 
                            })}
                            className="w-full p-1.5 border border-[#141414] text-[10px] font-mono focus:ring-1 focus:ring-[#141414] bg-white"
                          >
                            {headers.map(h => (
                              <option key={h} value={h}>{h}</option>
                            ))}
                          </select>
                        </div>
                      ))}
                    </div>
                  </div>
                  {/* Reset Button */}
                  <div className="flex items-end pb-1 lg:col-span-3">
                    <button 
                      onClick={() => updateConfig(DEFAULT_CONFIG)}
                      className="flex items-center gap-2 px-4 py-2 border border-[#141414] hover:bg-[#141414] hover:text-[#E4E3E0] transition-all text-[10px] font-bold uppercase tracking-tighter w-full justify-center"
                    >
                      <RotateCcw size={12} />
                      重置所有规则为默认参数
                    </button>
                  </div>
                </div>
              </motion.div>
            )}
          </AnimatePresence>
        </div>

        {/* Expense Type Filter */}
        {availableExpenseTypes.length > 0 && (
          <div className="mb-8 border border-[#141414] bg-white p-6">
            <div className="flex justify-between items-center mb-4">
              <div className="flex items-center gap-2">
                <Filter size={16} />
                <h2 className="font-serif italic text-sm uppercase tracking-wider">分析费用类型选择</h2>
                <select 
                  value={config.fieldMappings.expenseType}
                  onChange={handleExpenseTypeFieldChange}
                  className="ml-4 p-1 border border-[#141414] text-[11px] font-mono focus:ring-1 focus:ring-[#141414] bg-white"
                >
                  {headers.map(h => (
                    <option key={h} value={h}>{h}</option>
                  ))}
                </select>
              </div>
              <div className="flex gap-4">
                <button 
                  onClick={selectAllExpenseTypes}
                  className="text-[10px] font-mono uppercase opacity-50 hover:opacity-100 underline"
                >
                  全选
                </button>
                <button 
                  onClick={deselectAllExpenseTypes}
                  className="text-[10px] font-mono uppercase opacity-50 hover:opacity-100 underline"
                >
                  全不选
                </button>
              </div>
            </div>
            <div className="flex flex-wrap gap-3 max-h-64 overflow-y-auto p-2 border border-[#141414]/10">
              {availableExpenseTypes.slice(0, 100).map(type => (
                <button
                  key={type}
                  onClick={() => toggleExpenseType(type)}
                  className={cn(
                    "px-3 py-1 text-[11px] font-mono border transition-all",
                    selectedExpenseTypes.includes(type)
                      ? "bg-[#141414] text-[#E4E3E0] border-[#141414]"
                      : "bg-transparent text-[#141414] border-[#141414]/20 hover:border-[#141414]"
                  )}
                >
                  {type}
                </button>
              ))}
              {availableExpenseTypes.length > 100 && (
                <div className="text-[10px] font-mono opacity-50 flex items-center px-2">
                  ...及其他 {availableExpenseTypes.length - 100} 项 (已隐藏，可能选择了错误的字段)
                </div>
              )}
            </div>
          </div>
        )}

        {/* Action Button for Start Audit */}
        {records.length > 0 && (
          <div className="mb-8 flex flex-col items-center justify-center gap-3">
            <button
              onClick={() => runAudit(records, config, selectedExpenseTypes)}
              disabled={isProcessing}
              className={cn(
                "px-12 py-4 font-bold tracking-widest uppercase transition-all flex items-center gap-2 shadow-lg",
                isAuditOutdated 
                  ? "bg-[#F27D26] text-white hover:bg-[#d96b1c] scale-105 animate-pulse" 
                  : "bg-[#141414] text-white hover:bg-black disabled:opacity-50"
              )}
            >
              {isProcessing ? (
                <>
                  <RefreshCw className="animate-spin" size={20} />
                  正在审计...
                </>
              ) : (
                <>
                  <Play size={20} />
                  {isAuditOutdated ? '重新开始审计' : '开始审计'}
                </>
              )}
            </button>
            {isAuditOutdated && (
              <span className="text-xs text-[#F27D26] font-mono">
                * 配置已更改，请点击重新运行审计
              </span>
            )}
          </div>
        )}

        {/* Stats Grid */}
        <div className="grid grid-cols-1 md:grid-cols-4 gap-6 mb-8">
          {[
            { label: '总记录数', value: records.length, icon: Database },
            { label: '检测到风险', value: risks.length, icon: AlertTriangle, color: 'text-red-600' },
            { label: '高风险项', value: risks.filter(r => r.severity === 'high').length, icon: Shield, color: 'text-red-800' },
            { 
              label: '审计状态', 
              value: records.length === 0 ? '待机' : isAuditOutdated ? '待刷新' : '完成', 
              icon: isAuditOutdated ? RefreshCw : CheckCircle2, 
              color: records.length === 0 ? 'text-gray-400' : isAuditOutdated ? 'text-yellow-600' : 'text-green-600' 
            },
          ].map((stat, i) => (
            <motion.div 
              key={i}
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              transition={{ delay: i * 0.1 }}
              className="border border-[#141414] p-4 bg-white/50 backdrop-blur-sm"
            >
              <div className="flex justify-between items-start mb-2">
                <span className="text-[10px] font-mono uppercase opacity-50 tracking-widest">{stat.label}</span>
                <stat.icon size={14} className="opacity-30" />
              </div>
              <div className={cn("text-3xl font-bold tracking-tighter", stat.color)}>
                {stat.value}
              </div>
            </motion.div>
          ))}
        </div>

        {/* Risk Distribution */}
        {risks.length > 0 && (
          <div className="mb-8 border border-[#141414] bg-white p-6">
            <h2 className="font-serif italic text-sm uppercase tracking-wider mb-4 flex items-center gap-2">
              <AlertTriangle size={16} />
              命中风险规则分布
            </h2>
            <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
              {riskDistribution.map(([ruleName, count]) => (
                <div key={ruleName} className="border border-[#141414]/20 p-3 flex justify-between items-center bg-gray-50">
                  <span className="text-[11px] font-bold">{ruleName}</span>
                  <span className="text-xs font-mono bg-[#141414] text-[#E4E3E0] px-2 py-0.5 rounded-full">{count}</span>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* Main Content */}
        <div className="border border-[#141414] bg-white overflow-hidden">
          <div className="border-b border-[#141414] p-4 flex justify-between items-center bg-[#f5f5f5]">
            <div className="flex items-center gap-4">
              <h2 className="font-serif italic text-sm uppercase tracking-wider">风险分析列表</h2>
              {fileName && (
                <span className="text-[10px] font-mono bg-[#141414] text-[#E4E3E0] px-2 py-0.5 rounded-full">
                  {fileName}
                </span>
              )}
            </div>
            <div className="flex items-center gap-2">
              <div className="relative">
                <Search size={14} className="absolute left-3 top-1/2 -translate-y-1/2 opacity-30" />
                <input 
                  type="text" 
                  placeholder="搜索记录..." 
                  value={searchTerm}
                  onChange={(e) => setSearchTerm(e.target.value)}
                  className="pl-9 pr-4 py-1.5 border border-[#141414] text-[10px] font-mono uppercase focus:outline-none focus:ring-1 focus:ring-[#141414] w-64"
                />
              </div>
            </div>
          </div>

          <div className="overflow-x-auto">
            <table className="w-full text-left border-collapse">
              <thead>
                <tr className="border-b border-[#141414] bg-[#f9f9f9]">
                  <th className="p-4 text-[11px] font-serif italic opacity-50 uppercase tracking-widest border-r border-[#141414]">报销单编号</th>
                  <th className="p-4 text-[11px] font-serif italic opacity-50 uppercase tracking-widest border-r border-[#141414]">所属员工或机构</th>
                  <th className="p-4 text-[11px] font-serif italic opacity-50 uppercase tracking-widest border-r border-[#141414]">风险规则</th>
                  <th className="p-4 text-[11px] font-serif italic opacity-50 uppercase tracking-widest border-r border-[#141414]">严重程度</th>
                  <th className="p-4 text-[11px] font-serif italic opacity-50 uppercase tracking-widest">风险描述</th>
                </tr>
              </thead>
              <tbody>
                <AnimatePresence mode="popLayout">
                  {paginatedRisks.length > 0 ? (
                    paginatedRisks.map((risk, idx) => (
                      <motion.tr 
                        key={`${risk.recordId}-${idx}`}
                        initial={{ opacity: 0 }}
                        animate={{ opacity: 1 }}
                        exit={{ opacity: 0 }}
                        className="border-b border-[#141414]/10 hover:bg-[#141414] hover:text-[#E4E3E0] transition-colors group cursor-default"
                      >
                        <td className="p-4 font-mono text-[11px] border-r border-[#141414]/10">{risk.recordId}</td>
                        <td className="p-4 text-xs font-bold border-r border-[#141414]/10">
                          {(() => {
                            if (risk.entityName) return risk.entityName;
                            const rec = recordMap.get(String(risk.recordId));
                            return rec ? String(rec[config.fieldMappings.employee] || '未知') : '未知';
                          })()}
                        </td>
                        <td className="p-4 text-xs border-r border-[#141414]/10">
                          <span className="px-2 py-0.5 border border-current rounded-full text-[9px] font-bold uppercase">
                            {risk.ruleName}
                          </span>
                        </td>
                        <td className="p-4 border-r border-[#141414]/10">
                          <div className="flex items-center gap-2">
                            <div className={cn(
                              "w-2 h-2 rounded-full",
                              risk.severity === 'high' ? "bg-red-600" : 
                              risk.severity === 'medium' ? "bg-orange-500" : "bg-blue-500"
                            )} />
                            <span className="text-[10px] font-bold uppercase">
                              {risk.severity === 'high' ? '高' : risk.severity === 'medium' ? '中' : '低'}
                            </span>
                          </div>
                        </td>
                        <td className="p-4 text-xs opacity-80 group-hover:opacity-100">
                          {risk.description}
                        </td>
                      </motion.tr>
                    ))
                  ) : (
                    <tr>
                      <td colSpan={5} className="p-20 text-center">
                        <div className="flex flex-col items-center gap-4 opacity-20">
                          <Database size={48} />
                          <p className="font-serif italic text-lg">
                            {records.length > 0 ? "未找到匹配的风险记录。" : "未加载审计数据。请导入 Excel 文件。"}
                          </p>
                        </div>
                      </td>
                    </tr>
                  )}
                </AnimatePresence>
              </tbody>
            </table>
          </div>
          
          {totalPages > 1 && (
            <div className="flex items-center justify-between mt-4 px-2">
              <div className="text-xs font-mono opacity-50">
                显示 {(currentPage - 1) * itemsPerPage + 1} - {Math.min(currentPage * itemsPerPage, filteredRisks.length)} 条，共 {filteredRisks.length} 条
              </div>
              <div className="flex items-center gap-2">
                <button
                  onClick={() => setCurrentPage(p => Math.max(1, p - 1))}
                  disabled={currentPage === 1}
                  className="px-3 py-1 border border-[#141414] text-xs font-mono hover:bg-[#141414] hover:text-[#E4E3E0] disabled:opacity-30 disabled:hover:bg-transparent disabled:hover:text-[#141414] transition-colors"
                >
                  上一页
                </button>
                <span className="text-xs font-mono px-2">
                  {currentPage} / {totalPages}
                </span>
                <button
                  onClick={() => setCurrentPage(p => Math.min(totalPages, p + 1))}
                  disabled={currentPage === totalPages}
                  className="px-3 py-1 border border-[#141414] text-xs font-mono hover:bg-[#141414] hover:text-[#E4E3E0] disabled:opacity-30 disabled:hover:bg-transparent disabled:hover:text-[#141414] transition-colors"
                >
                  下一页
                </button>
              </div>
            </div>
          )}
        </div>

        {/* Footer Info */}
        <footer className="mt-12 pt-6 border-t border-[#141414]/20 flex justify-between items-center">
          <div className="flex gap-8">
            <div className="flex items-center gap-2">
              <Info size={14} className="opacity-30" />
              <span className="text-[10px] font-mono opacity-50 uppercase tracking-widest">合规引擎: v4.2.0-stable</span>
            </div>
            <div className="flex items-center gap-2">
              <Database size={14} className="opacity-30" />
              <span className="text-[10px] font-mono opacity-50 uppercase tracking-widest">数据完整性: 已验证</span>
            </div>
          </div>
          <div className="text-[10px] font-mono opacity-30 uppercase tracking-widest">
            © 2026 EagleEye Systems. 保留所有权利。
          </div>
        </footer>
      </main>

      {/* Processing Overlay */}
      {isProcessing && (
        <div className="fixed inset-0 bg-[#E4E3E0]/80 backdrop-blur-md z-50 flex items-center justify-center">
          <div className="flex flex-col items-center gap-4">
            <div className="w-12 h-12 border-4 border-[#141414] border-t-transparent rounded-full animate-spin" />
            <p className="font-mono text-xs font-bold uppercase tracking-widest animate-pulse">正在分析记录...</p>
          </div>
        </div>
      )}
    </div>
  );
}
