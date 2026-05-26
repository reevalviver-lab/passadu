แก้ไขบัค blank page และข้อมูลไม่แสดง

ไฟล์ที่แก้หลัก:
- Code.gs: แก้ doGet() ให้ใช้ HtmlService.createTemplateFromFile('Index').evaluate() เพื่อให้ include ทำงานถูกต้อง
- Index.html: กลับมาใช้โครงสร้าง template หลัก และ include ไฟล์ย่อยอย่างถูกต้อง
- JavaScript.html: เพิ่ม/ซ่อม renderItems(), openCategoryForm(), saveCategorySubmit(), deleteCategory(); ปรับ bindForms() ให้ไม่ทำให้สคริปต์หยุดเมื่อบางฟอร์มไม่มีอยู่; เพิ่ม fallback ให้ APP.categories/APP.items เป็น array; ป้องกัน form submit แล้ว reload หน้า

วิธีใช้:
1) แทนที่ไฟล์ทั้งหมดใน Google Apps Script ด้วยไฟล์ใน ZIP นี้
2) บันทึก แล้ว Deploy > New deployment หรือ Manage deployments > Edit > New version
3) รีเฟรช Web App ใหม่
หกดกหดกหด
