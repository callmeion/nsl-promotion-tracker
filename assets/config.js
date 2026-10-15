// ตั้งค่าระบบ — แก้ไขที่ไฟล์นี้
window.APP_CONFIG = {
  dataFile: 'data/promotion.xlsx',   // ไฟล์ Excel ใน repo เดียวกัน (โหลดอัตโนมัติ)

  // รหัสผ่าน = SHA-256 ของรหัสที่ต้องการ (ค่าปัจจุบันคือ N@sl2026)
  // วิธีเปลี่ยน: เปิด Console (F12) พิมพ์
  //   crypto.subtle.digest('SHA-256', new TextEncoder().encode('รหัสใหม่')).then(b=>console.log([...new Uint8Array(b)].map(x=>x.toString(16).padStart(2,'0')).join('')))
  // แล้วนำค่าที่ได้มาวางแทนด้านล่าง
  passwordHash: '1502610a978d687e50d270cbf9147eae0093a1202722e9db1ebdb2f208662261',

  // สถานะที่ใช้ (ต้องตรงกับคอลัมน์ V ใน Excel) และสีที่แสดง
  statuses: [
    { key: 'ยังไม่สัมภาษณ์',      color: 'gray'  },
    { key: 'นัดสัมภาษณ์แล้ว',     color: 'blue'  },
    { key: 'สัมภาษณ์แล้ว (รอผล)', color: 'amber' },
    { key: 'ผ่าน',               color: 'green' },
    { key: 'ไม่ผ่าน',             color: 'red'   }
  ]
};
