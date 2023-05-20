# Python-Excel
Giới thiệu về sử dụng Python trong Excel

## Xin chào
   Xin chào toàn thể ACE.
   Theo đề nghị của add, và có thể 1 vài ACE cũng quan tâm về việc sử dụng Python trong Excel, mình xin giới thiệu đôi chút về vấn đề này nhé. ^^
   Ở đây, có thể có lớn, bằng hoặc nhỏ tuổi hơn mình, vì vậy cho mình xưng tên nhé
   Dũng xin chào mọi người

## Giới thiệu
   Hiện tại, Python có khá nhiều thư viện để làm việc với Excel, có những công cụ rất hay (tạo user define functions cho Excel có ghi chú, tạo dynamic array (hay Function cho kết quả như Sub)...)
   <br/>
   Một trong số đó là *Xlwings*
   <br/>
       + Scripting: Tự động hóa/tương tác với Excel từ môi trường Python, sử dụng cú pháp gần với VBA mà vẫn `Pythonic`.
       <br/>
       + Macros: Viết các script python thay thế cho VBA macros, giúp code dễ đọc hơn. Sau khi viết script python xong, chỉ cần gọi 1 hàm trong VBA là script chạy.
       <br/>
       + UDFs: Viết hàm người dùng tự định nghĩa bằng ngôn ngữ Python và sử dụng được hàm đó trong excel (Windows only).
       <br/>
       + REST API: cung cấp REST API cho Excel workbook.
   <br/>
   Tham khảo thêm: https://docs.xlwings.org/en/latest/index.html


## Cài đặt thư viện
Ở phạm vi các bài viết ở đây sẽ áp dụng với máy tính sử dụng hệ điều hành Windows và sử dụng thư viện openpyxl

- Cài đặt phần mềm Python (hoàn toàn miễn phí, dung lượng nhỏ) lên máy tính: https://www.python.org/downloads/
   <br/><br/>
   Chọn phiên bản phù hợp (x86, x64) với máy tính

- Thư viện Python:
   + pip install pandas
   + pip install numpy
   + pip install openpyxl
   + pip install xlwings
   + xlwings addin install
<br/><br/>
*Cách cài đặt: Sau khi cài đặt thành công Python ở bước trên thì mở Command Prompt của Windows lên và gõ lệnh sau rồi nhấn enter*

- Tạo Folder named `MyLove`

- Add vào excel file named 'Nai' và lưu dạng xlsm (Tên tùy chọn ace ha. Nai là tên đẹp nhất thôi, kk)

- Add vào file 'Nai.py' (file Python cùng tên File Excel để đỡ setup)

- Enable Trust access to the VBA project object model under (chỉ 1 lần thoy): `File > Options > Trust Center > Trust Center Settings > Macro Settings`

- Thêm code vào file .file như trong repo. Các bạn có thể điều chỉnh code nếu muốn

- Ở file excel, sang tab xlwings và chọn *Import Function* để add code Python vào. Thay đổi code ở file, các bạn phải nhấn để load lại

- Gõ lệnh Script hàm đã đặt tên để có kết quả

      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
     
