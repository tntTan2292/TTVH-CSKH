import React, { useState, useEffect, useMemo } from 'react';
import axios from 'axios';
import { 
  Radar, RadarChart, PolarGrid, PolarAngleAxis, PolarRadiusAxis, 
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, 
  ResponsiveContainer, Cell
} from 'recharts';
import { 
  Layout, Card, Row, Col, Statistic, Table, Tag, 
  Button, Typography, Modal, Input, message, Badge, Space, 
  Progress, Divider, Tabs, Empty, Spin, Upload
} from 'antd';
import { 
  DashboardOutlined, AlertOutlined, FileExcelOutlined, 
  CheckCircleOutlined, ClockCircleOutlined, RocketOutlined, 
  SafetyCertificateOutlined, UserOutlined, LockOutlined, 
  LogoutOutlined, EnvironmentOutlined, ArrowUpOutlined, 
  SearchOutlined, WarningOutlined, SyncOutlined, DeleteOutlined,
  CloudSyncOutlined, LoadingOutlined, ImportOutlined,
  UploadOutlined
} from '@ant-design/icons';

const { Header, Content } = Layout;
const { Title, Text } = Typography;

const VNPOST_NAVY = '#00387b';
const VNPOST_GOLD = '#fdb913';

// Set base URL for API
const API_BASE = 'http://localhost:8000';

const App = () => {
  const [data, setData] = useState(null);
  const [loading, setLoading] = useState(true);
  const [uploading, setUploading] = useState(false);

  const fetchData = async () => {
    setLoading(true);
    try {
      const res = await axios.get(`${API_BASE}/stats`);
      setData(res.data);
    } catch (err) {
      console.error("Fetch Error:", err);
      message.error("Không thể kết nối đến máy chủ dữ liệu.");
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchData();
  }, []);

  const handleUpload = async (file) => {
    const formData = new FormData();
    formData.append('file', file);
    
    setUploading(true);
    try {
      const res = await axios.post(`${API_BASE}/upload`, formData);
      message.success(res.data.message);
      fetchData(); // Refresh data
    } catch (err) {
      message.error("Lỗi nạp file: " + (err.response?.data?.detail || err.message));
    } finally {
      setUploading(false);
    }
    return false; // Prevent default upload behavior
  };

  if (loading && !data) {
    return <div className="h-screen w-screen flex items-center justify-center bg-gray-100"><Spin size="large" tip="Đang tải dữ liệu VIP..." /></div>;
  }

  const { customer, kpis, radarData, directionData, bottlenecks, slaList } = data || {
    customer: { name: 'Chưa có dữ liệu' },
    kpis: { total: 0, success: 0, pending: 0, sla: 0 },
    radarData: [],
    directionData: [],
    bottlenecks: [],
    slaList: []
  };

  return (
    <Layout style={{ minHeight: '100vh', background: '#f0f2f5' }}>
      <Header className="vnpost-header">
        <div className="flex items-center">
          <RocketOutlined style={{ fontSize: 32, color: VNPOST_GOLD, marginRight: 12 }} />
          <div className="flex flex-col">
            <Title level={3} style={{ color: '#fff', margin: 0, lineHeight: 1 }}>VNPOST HUẾ</Title>
            <Text style={{ color: VNPOST_GOLD, fontSize: 10, fontWeight: 900 }}>VIP CUSTOMER SERVICE DASHBOARD</Text>
          </div>
        </div>
        
        <Space size="middle">
          <div className="text-white mr-4 flex flex-col items-end">
            <Text style={{ color: '#fff', fontSize: 12 }}>Đang theo dõi VIP:</Text>
            <Text strong style={{ color: VNPOST_GOLD, fontSize: 14 }}>{customer.name}</Text>
          </div>
          
          <Upload beforeUpload={handleUpload} showUploadList={false}>
            <Button 
              type="primary" 
              icon={<ImportOutlined />} 
              style={{ background: '#52c41a', border: 'none', fontWeight: 'bold' }}
              loading={uploading}
            >
              NHẬP EXCEL MỚI
            </Button>
          </Upload>
          
          <Button icon={<SyncOutlined />} onClick={fetchData} loading={loading}>LÀM MỚI</Button>
        </Space>
      </Header>

      <Content style={{ padding: '24px' }}>
        {/* KPI CARDS */}
        <Row gutter={[16, 16]} className="mb-6">
          <Col xs={24} sm={12} md={6}>
            <Card className="kpi-card">
              <Statistic title={<Text strong>TỔNG SẢN LƯỢNG VIP</Text>} value={kpis.total} prefix={<DashboardOutlined style={{color: VNPOST_NAVY}} />} />
              <Progress percent={100} strokeColor={VNPOST_NAVY} size="small" showInfo={false} />
            </Card>
          </Col>
          <Col xs={24} sm={12} md={6}>
            <Card className="kpi-card">
              <Statistic title={<Text strong style={{color: '#52c41a'}}>PHÁT THÀNH CÔNG</Text>} value={kpis.success} valueStyle={{color: '#52c41a'}} prefix={<CheckCircleOutlined />} />
              <div className="flex items-center justify-between">
                <Progress percent={kpis.total ? Math.round((kpis.success/kpis.total)*100) : 0} strokeColor="#52c41a" size="small" style={{width: '70%'}} />
                <Badge count={kpis.success} overflowCount={999} style={{ backgroundColor: '#52c41a' }} />
              </div>
              <div className="mt-2 text-xs text-gray-400 italic">Mốc 30 đơn: {kpis.success}/30</div>
              <Progress percent={Math.min(100, Math.round((kpis.success/30)*100))} strokeColor={VNPOST_GOLD} status="active" />
            </Card>
          </Col>
          <Col xs={24} sm={12} md={6}>
            <Card className="kpi-card">
              <Statistic title={<Text strong style={{color: VNPOST_GOLD}}>CHỜ PHÁT / LỖI</Text>} value={kpis.pending} valueStyle={{color: VNPOST_GOLD}} prefix={<ClockCircleOutlined />} />
              <Progress percent={kpis.total ? Math.round((kpis.pending/kpis.total)*100) : 0} strokeColor={VNPOST_GOLD} size="small" />
            </Card>
          </Col>
          <Col xs={24} sm={12} md={6}>
            <Card className={`kpi-card ${kpis.sla > 0 ? 'sla-critical' : ''}`}>
              <Statistic title={<Text strong style={{color: '#ff4d4f'}}>VI PHẠM SLA (3 NGÀY)</Text>} value={kpis.sla} valueStyle={{color: '#ff4d4f', fontWeight: 'bold'}} prefix={<AlertOutlined />} />
              <Tag color="red" icon={<WarningOutlined />}>Báo cáo lãnh đạo ngay</Tag>
            </Card>
          </Col>
        </Row>

        {/* CHARTS */}
        <Row gutter={[16, 16]} className="mb-6">
          <Col xs={24} lg={12}>
            <Card title={<><EnvironmentOutlined /> Hiệu suất phát VIP theo Tỉnh</>} className="kpi-card">
              <div style={{ height: 350 }}>
                {radarData.length > 0 ? (
                  <ResponsiveContainer width="100%" height="100%">
                    <RadarChart data={radarData}>
                      <PolarGrid />
                      <PolarAngleAxis dataKey="subject" tick={{fontSize: 10, fontWeight: 600}} />
                      <Radar name="Sản lượng" dataKey="A" stroke={VNPOST_NAVY} fill={VNPOST_NAVY} fillOpacity={0.6} />
                      <Radar name="Hiệu suất (%)" dataKey="B" stroke={VNPOST_GOLD} fill={VNPOST_GOLD} fillOpacity={0.5} />
                      <Tooltip />
                    </RadarChart>
                  </ResponsiveContainer>
                ) : <div className="flex h-full items-center justify-center"><Empty description="Chưa có dữ liệu" /></div>}
              </div>
            </Card>
          </Col>
          <Col xs={24} lg={12}>
            <Card title={<><ArrowUpOutlined /> Phân bổ theo Hướng Đóng chuyển</>} className="kpi-card">
              <div style={{ height: 350 }}>
                {directionData.length > 0 ? (
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart data={directionData}>
                      <CartesianGrid strokeDasharray="3 3" vertical={false} />
                      <XAxis dataKey="name" />
                      <YAxis />
                      <Tooltip cursor={{fill: 'transparent'}} />
                      <Bar dataKey="value" name="Bưu gửi" fill={VNPOST_NAVY} radius={[4, 4, 0, 0]}>
                         {directionData.map((entry, index) => (
                            <Cell key={`cell-${index}`} fill={index % 2 === 0 ? VNPOST_NAVY : VNPOST_GOLD} />
                          ))}
                      </Bar>
                    </BarChart>
                  </ResponsiveContainer>
                ) : <div className="flex h-full items-center justify-center"><Empty description="Chưa có dữ liệu" /></div>}
              </div>
            </Card>
          </Col>
        </Row>

        {/* BOTTOM SECTION */}
        <Row gutter={[16, 16]}>
          <Col xs={24} xl={10}>
            <Card title={<><WarningOutlined style={{color: '#ff4d4f'}} /> Điểm nghẽn tại Bưu cục (Top 10 Tồn)</>} className="kpi-card">
              <Table 
                dataSource={bottlenecks} 
                pagination={false} 
                size="small" 
                rowKey="name"
                columns={[
                  { title: 'Bưu cục (BCVH)', dataIndex: 'name', render: t => <Text strong>{t}</Text> },
                  { title: 'Tồn phát', dataIndex: 'backlog', render: v => <Badge count={v} style={{backgroundColor: v > 5 ? '#f5222d' : '#1890ff'}} /> },
                  { title: 'Quá SLA', dataIndex: 'sla', render: v => <Tag color={v > 0 ? "red" : "green"}>{v}</Tag> }
                ]}
              />
            </Card>
          </Col>
          <Col xs={24} xl={14}>
            <Card title={<><SearchOutlined /> Action Center: Danh sách đơn VIP lỗi SLA</>} className="kpi-card">
              <Table 
                dataSource={slaList} 
                rowKey="id" 
                size="small" 
                pagination={{pageSize: 8}}
                columns={[
                  { title: 'Số hiệu', dataIndex: 'id', width: 140 },
                  { title: 'Aging', dataIndex: 'aging', render: v => <Tag color="red" className="font-bold">{v} ngày</Tag> },
                  { title: 'BCVH', dataIndex: 'bcvh' },
                  { title: 'Vị trí cuối', dataIndex: 'lastPos', ellipsis: true },
                ]}
              />
            </Card>
          </Col>
        </Row>
      </Content>
    </Layout>
  );
};

export default App;
