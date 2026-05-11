import React, { useState, useEffect, Component } from 'react';
import axios from 'axios';
import { 
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip as ReTooltip, 
  ResponsiveContainer, Cell
} from 'recharts';
import { 
  Layout, Row, Col, Table, Tag, Progress, Button, message, Space, 
  Typography, Badge, Modal, Result, Empty, Spin
} from 'antd';
import { 
  DashboardOutlined, AlertOutlined, CheckCircleOutlined, 
  ClockCircleOutlined, RocketOutlined, SyncOutlined, 
  WarningOutlined, ImportOutlined, EnvironmentOutlined
} from '@ant-design/icons';

const { Header, Content } = Layout;
const { Title, Text } = Typography;

const API_BASE = 'http://localhost:8010/api/dashboard';
const UPLOAD_URL = 'http://localhost:8010/upload';

// THEME TOKENS
const COLORS = {
  bg: '#000c17',
  card: '#001529',
  accent: '#fadb14',
  success: '#52c41a',
  danger: '#ff4d4f'
};

class ErrorBoundary extends Component {
  constructor(props) { super(props); this.state = { hasError: false }; }
  static getDerivedStateFromError(error) { return { hasError: true }; }
  render() {
    if (this.state.hasError) return <div style={{padding: 20, color: '#ff4d4f', background: COLORS.card, borderRadius: 8}}>⚠️ Widget Crash - Vui lòng F5</div>;
    return this.props.children;
  }
}

const App = () => {
  const [stats, setStats] = useState(null);
  const [bcvhData, setBcvhData] = useState([]);
  const [bottlenecks, setBottlenecks] = useState([]);
  const [slaRisk, setSlaRisk] = useState([]);
  const [provinceData, setProvinceData] = useState([]);
  const [loading, setLoading] = useState(true);

  const fetchAll = async () => {
    setLoading(true);
    try {
      const [s, b, n, r, p] = await Promise.all([
        axios.get(`${API_BASE}/stats`).then(res => res.data),
        axios.get(`${API_BASE}/bcvh-summary`).then(res => res.data),
        axios.get(`${API_BASE}/bcvh-bottleneck`).then(res => res.data),
        axios.get(`${API_BASE}/sla-risk`).then(res => res.data),
        axios.get(`${API_BASE}/province-performance`).then(res => res.data)
      ]);
      
      setStats(s && !s.error ? s : null);
      setBcvhData(Array.isArray(b) ? b : []);
      setBottlenecks(Array.isArray(n) ? n : []);
      setSlaRisk(Array.isArray(r) ? r : []);
      setProvinceData(Array.isArray(p) ? p : []);
    } catch (err) {
      console.error("API Failure:", err);
      message.error("Mất kết nối Enterprise Backend 8010");
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => { fetchAll(); }, []);

  const handleUpload = async () => {
    const input = document.createElement('input');
    input.type = 'file'; input.accept = '.xlsx';
    input.onchange = async (e) => {
      const formData = new FormData();
      formData.append('file', e.target.files[0]);
      const hide = message.loading('Đang khởi tạo Snapshot...', 0);
      try {
        const res = await axios.post(UPLOAD_URL, formData);
        message.success("Snapshot Created!");
        fetchAll();
      } catch (err) {
        Modal.error({ title: 'Import Error', content: err.response?.data?.detail || "Upload Failed" });
      } finally { hide(); }
    };
    input.click();
  };

  if (loading && !stats) return <div className="h-screen w-screen flex flex-col items-center justify-center" style={{background: COLORS.bg, color: 'white'}}><Spin size="large" /><p className="mt-4">Đang đồng bộ EXECUTIVE V2.0...</p></div>;

  if (!stats) return (
    <Layout style={{ minHeight: '100vh', background: COLORS.bg }}>
      <Header className="flex justify-between items-center" style={{ background: COLORS.card, borderBottom: `2px solid ${COLORS.accent}` }}>
        <Title level={4} style={{ color: 'white', margin: 0 }}>VNPOST HUE DOC</Title>
      </Header>
      <div className="flex-1 flex items-center justify-center">
        <Result status="info" title={<span style={{color: 'white'}}>System Ready</span>} subTitle={<span style={{color: 'rgba(255,255,255,0.5)'}}>Chưa có Snapshot. Vui lòng nạp dữ liệu.</span>} extra={<Button type="primary" size="large" onClick={handleUpload} style={{background: COLORS.accent, color: 'black', border: 'none'}}>BẮT ĐẦU IMPORT</Button>} />
      </div>
    </Layout>
  );

  const successRate = stats ? ((stats.kpis.success / stats.kpis.total) * 100).toFixed(1) : "0.0";

  return (
    <ErrorBoundary>
      <Layout style={{ minHeight: '100vh', background: COLORS.bg, color: 'white' }}>
        <Header className="flex justify-between items-center px-6" style={{ background: COLORS.card, borderBottom: `2px solid ${COLORS.accent}`, position: 'sticky', top: 0, z-index: 1000 }}>
          <Space size="large">
            <Title level={4} style={{ color: 'white', margin: 0, letterSpacing: 1 }}>
              VNPOST HUE <span style={{ color: COLORS.accent, fontSize: 12 }}>DOC</span>
              <span style={{ background: COLORS.success, color: 'black', padding: '2px 8px', borderRadius: 4, fontSize: 10, marginLeft: 12, fontWeight: 900 }}>V2.0 - PROVINCE MODE</span>
            </Title>
            <Tag color="gold" style={{ background: 'transparent', border: `1px solid ${COLORS.accent}`, color: COLORS.accent }}>SNAPSHOT V{stats.session_info.id}</Tag>
          </Space>
          <Space>
            <Button type="primary" onClick={handleUpload} style={{ background: COLORS.accent, color: 'black', fontWeight: 600, border: 'none' }} icon={<ImportOutlined />}>IMPORT DATA</Button>
            <Button icon={<SyncOutlined />} onClick={fetchAll} ghost />
          </Space>
        </Header>

        <Content className="p-6 max-w-[1920px] mx-auto w-full">
          {/* KPI CARDS */}
          <Row gutter={[16, 16]} className="mb-4">
            <Col span={6}>
              <div className="p-4 rounded-lg" style={{ background: COLORS.card, border: '1px solid rgba(255,255,255,0.1)' }}>
                <div style={{ color: 'rgba(255,255,255,0.5)', fontSize: 11, textTransform: 'uppercase', marginBottom: 8 }}>Tổng sản lượng</div>
                <div style={{ fontSize: 32, fontWeight: 700, color: COLORS.accent }}>{stats.kpis.total.toLocaleString()}</div>
              </div>
            </Col>
            <Col span={6}>
              <div className="p-4 rounded-lg" style={{ background: COLORS.card, border: '1px solid rgba(255,255,255,0.1)', borderLeft: `4px solid ${COLORS.success}` }}>
                <div style={{ color: 'rgba(255,255,255,0.5)', fontSize: 11, textTransform: 'uppercase', marginBottom: 8 }}>Phát thành công</div>
                <div style={{ fontSize: 32, fontWeight: 700, color: COLORS.success }}>{stats.kpis.success.toLocaleString()}</div>
                <div style={{ fontSize: 14, color: COLORS.success, fontWeight: 600 }}>{successRate}% Success Rate</div>
                <Progress percent={parseFloat(successRate)} size="small" strokeColor={COLORS.success} trailColor="rgba(255,255,255,0.1)" showInfo={false} className="mt-2" />
              </div>
            </Col>
            <Col span={6}>
              <div className="p-4 rounded-lg" style={{ background: COLORS.card, border: '1px solid rgba(255,255,255,0.1)' }}>
                <div style={{ color: 'rgba(255,255,255,0.5)', fontSize: 11, textTransform: 'uppercase', marginBottom: 8 }}>Tồn đang xử lý</div>
                <div style={{ fontSize: 32, fontWeight: 700, color: '#faad14' }}>{stats.kpis.pending.toLocaleString()}</div>
              </div>
            </Col>
            <Col span={6}>
              <div className="p-4 rounded-lg" style={{ background: COLORS.card, border: `1px solid ${COLORS.danger}` }}>
                <div style={{ color: 'rgba(255,255,255,0.5)', fontSize: 11, textTransform: 'uppercase', marginBottom: 8 }}>Vi phạm SLA</div>
                <div style={{ fontSize: 32, fontWeight: 700, color: COLORS.danger }}>{stats.kpis.sla.toLocaleString()}</div>
              </div>
            </Col>
          </Row>

          <Row gutter={[16, 16]}>
            <Col span={14}>
              <div className="p-4 rounded-lg" style={{ background: COLORS.card, border: '1px solid rgba(255,255,255,0.1)' }}>
                <Title level={5} style={{ color: 'white', marginBottom: 16 }}>📊 Hiệu suất Bưu cục Vận hành (BCVH)</Title>
                <Table 
                  dataSource={bcvhData} 
                  size="small"
                  pagination={{ pageSize: 8 }}
                  columns={[
                    { title: 'BƯU CỤC', dataIndex: 'name' },
                    { title: 'TỔNG', dataIndex: 'total', align: 'right' },
                    { title: 'THÀNH CÔNG', dataIndex: 'success', align: 'right', render: v => <span style={{color: '#b7eb8f'}}>{v}</span> },
                    { title: 'SLA', dataIndex: 'sla', align: 'right', render: v => <span style={{color: v > 0 ? COLORS.danger : 'inherit'}}>{v}</span> },
                    { title: 'TỶ LỆ (%)', dataIndex: 'rate', align: 'center', render: v => <Tag color={v > 90 ? 'green' : 'orange'}>{v}%</Tag> }
                  ]}
                  rowClassName="hover:bg-white/5"
                />
              </div>
            </Col>

            <Col span={10}>
              <div className="p-4 rounded-lg" style={{ background: COLORS.card, border: '1px solid rgba(255,255,255,0.1)' }}>
                <Title level={5} style={{ color: 'white', marginBottom: 16 }}>📈 Hiệu suất theo Tỉnh phát</Title>
                <div style={{ height: 330 }}>
                  {provinceData.length > 0 ? (
                    <ResponsiveContainer width="100%" height="100%">
                      <BarChart data={provinceData} layout="vertical" margin={{ left: 50, right: 30 }}>
                        <XAxis type="number" hide />
                        <YAxis dataKey="province" type="category" tick={{fill: 'white', fontSize: 10}} width={120} />
                        <ReTooltip cursor={{fill: 'rgba(255,255,255,0.05)'}} contentStyle={{background: COLORS.card, border: `1px solid ${COLORS.accent}`, color: 'white'}} />
                        <Bar dataKey="rate" radius={[0, 4, 4, 0]}>
                          {provinceData.map((entry, index) => (
                            <Cell key={`cell-${index}`} fill={entry.rate > 85 ? COLORS.success : '#faad14'} />
                          ))}
                        </Bar>
                      </BarChart>
                    </ResponsiveContainer>
                  ) : <div className="h-full flex items-center justify-center text-gray-500">Không có dữ liệu tỉnh phát</div>}
                </div>
              </div>
            </Col>

            <Col span={24}>
              <div className="p-4 rounded-lg" style={{ background: COLORS.card, border: '1px solid rgba(255,255,255,0.1)' }}>
                <Title level={5} style={{ color: 'white', marginBottom: 16 }}>🛡️ Action Center (SLA Risk Management)</Title>
                <Table 
                  dataSource={slaRisk} 
                  size="small"
                  pagination={{ pageSize: 5 }}
                  columns={[
                    { title: 'MÃ BƯU GỬI', dataIndex: 'tracking_id' },
                    { title: 'TỈNH PHÁT', dataIndex: 'province' },
                    { title: 'BƯU CỤC VẬN HÀNH', dataIndex: 'post_office_name' },
                    { title: 'TUỔI ĐƠN', dataIndex: 'aging', render: v => <Tag color="error">{v} NGÀY</Tag> },
                    { title: 'HÀNH ĐỘNG', render: () => <Button size="small" ghost type="primary">Xử lý ngay</Button> }
                  ]}
                />
              </div>
            </Col>
          </Row>
        </Content>

        <style>{`
          .ant-table { background: transparent !important; color: white !important; }
          .ant-table-thead > tr > th { background: rgba(255,255,255,0.05) !important; color: rgba(255,255,255,0.6) !important; border-bottom: 1px solid rgba(255,255,255,0.1) !important; }
          .ant-table-tbody > tr > td { border-bottom: 1px solid rgba(255,255,255,0.05) !important; }
          .ant-table-cell { color: white !important; }
          .ant-pagination-item a { color: white !important; }
        `}</style>
      </Layout>
    </ErrorBoundary>
  );
};

export default App;
