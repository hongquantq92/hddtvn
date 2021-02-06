6 Hàm Excel thông dụng nhất kế toán cần nắm được
================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/6-ham-excel-thong-dung-nhat-ke-toan-can-nam-duoc/)Excel là phần mềm không còn xa lạ gì với dân kế toán. Nhưng không hẳn kế toán nào cũng thông thạo các thủ thuật tin học văn phòng trong công việc. Hiện nay đã có nhiều phần mềm online tiện lợi nhưng đôi khi kế toán vẫn phải dùng đến Excel. Dưới đây là …
-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Excel là phần mềm không còn xa lạ gì với dân kế toán. Nhưng không hẳn kế toán nào cũng thông thạo các thủ thuật tin học văn phòng trong công việc. Hiện nay đã có nhiều phần mềm online tiện lợi nhưng đôi khi kế toán vẫn phải dùng đến Excel. Dưới đây là 6 hàm Excel thông dụng trong kế toán mà bạn nên nắm rõ.**


![Hình ảnh có liên quan](https://hddtvn.com/wp-content/uploads/2021/01/20-ham-excel-thuong-dung-trong-ke-toan-01-99.jpg)


1. Hàm excel COUNTIF
--------------------


COUNTIF là hàm có điều kiện, cú pháp: **=COUNTIF(Vùng điều kiện; điều kiện đếm)**


Các bạn kế toán thường sử dụng hàm này khi lập bảng chấm công. Khi cần đếm số công của một nhân viên trong tháng thì dùng hàm này cho ra kết quả rất nhanh, cú pháp dễ nhớ.


Ví dụ, ký hiệu chấm công 1 ngày là dấu trừ (-), bạn cần tính xem từ ngày 1 – 31 nhân viên A làm bao nhiêu ngày, bạn dùng công thức:


=COUNTIF(D10:AH10;”-“) => kết quả trả về sẽ là tổng ký tự (-) trong vùng.


[**>>> Mách bạn cách sử dụng hàm COUNTIF trong Excel**](#)


2. Hàm excel SUMIF
------------------


SUMIF là hàm tính tổng theo điều kiện, công thức: **=SUMIF(Dãy ô điều kiện; Điều kiện cần tính; Dãy ô tính tổng)**


– Dãy ô điều kiện: Là dãy ô chứa điều kiện cần tính. Trong kế toán thường là Cột TK Nợ, TK Có, Cột chứa các mã hàng hóa, mã TK…


– Điều kiện cần tính: Chính là việc bạn đi tính tổng cho “cái gì”.


– Dãy ô tính tổng: Là dãy ô chứa giá trị tương ứng (số tiền, số lượng…) liên quan đến điều kiện cần tính. Trong kế toán có thể là dãy ô phát sinh Nợ hoặc Có, hoặc dãy ô chứa số lượng nhập, xuất…


Kế toán thường dùng hàm này khi:


– Tổng hợp số liệu từ sổ NKC lên Phát sinh Nợ Phát sinh Có trên Bảng cân đối số phát sinh Tài khoản.


– Tổng hợp số liệu từ PNK, PXK lên “Bảng nhập xuất tồn”.


– Tổng hợp số liệu từ sổ NKC lên cột PS Nợ, PS Có của “Bảng tổng hợp phải thu, phải trả khách hàng”.


– Kết chuyển bút toán cuối kỳ.


**Ví dụ:**  = SUMIF(B15:B40,”>=10;<=20″)


– Tính tổng của các giá trị trong vùng từ B15 đến B40 với điều kiện là các giá trị lớn hơn hoặc bằng 10 và nhỏ hơn hoặc bằng 20.


[**>>> Cách sử dụng hàm SUMIF trong Excel**](#)


3. Hàm excel IF
---------------


Đây được coi là hàm excel thông dụng, dễ nhất trong các hàm excel, công thức: **=IF(Điều kiện; Giá trị 1; Giá trị 2)**


Hàm IF trả về “Giá trị 1” nếu điều kiện đúng và nếu sai trả về “Giá trị 2” nếu điều kiện sai.


Ví dụ: =IF(B20=115;”MUA”;”BÁN”)


![Kết quả hình ảnh cho hàm excel thường dùng trong kế toán](https://hddtvn.com/wp-content/uploads/2021/01/cac-ham-trong-excel.jpg)


**[>>>Cách dùng hàm IF và cách kết hợp với các hàm khác trong Excel](#)**


4. Hàm excel VLOOKUP
--------------------


Hàm Vlookup là hàm trả về giá trị dò tìm theo cột đưa từ bảng tham chiếu lên bảng cơ sở dữ liệu theo đúng giá trị dò tìm.


Công thức: **=VLOOKUP(giá trị dò, bảng dò, cột giá trị trả về, kiểu dò)**


– Giá trị dò tìm: phải có Tên trong vùng dữ liệu tìm kiếm.


– Vùng dữ liệu tìm kiếm: phải chứa tên của “Giá trị dò tìm” và phải chứa “Giá trị cần tìm”. Điểm bắt đầu của vùng được tính từ dãy ô có chứa “giá trị dò tìm”.


– Cột trả về giá trị tìm kiếm: số thứ tự cột, tính từ bên trái sang của vùng dữ liệu tìm kiếm.


– Tham số “N”: N=0: dò tìm tuyệt đối (thường sử dụng); N=1: dò tìm tương đối.


Kế toán thường dùng hàm VLOOKUP khi:


– Tìm đơn giá Xuất kho từ bên Bảng nhập xuất tồn về Phiếu xuất kho.


– Tìm tên hàng hoá từ Danh mục hàng hoá về Bảng nhập xuất tồn.


– Tìm tên TK, mã TK từ Danh mục tài khoản.


– Tìm số Khấu hao (Phân bổ) luỹ kế từ kỳ trước, căn cứ vào Giá trị khấu hao( phân bổ) luỹ kế (của bảng , 242, 214).


[**>>>Hướng dẫn dùng hàm VLOOKUP trong Excel**](#)


5. Hàm excel LEFT
-----------------


LEFT là hàm lọc ký tự bên trái của chuỗi, công thức: **=LEFT(Chuỗi;N)**


– Chuỗi: dãy ký tự trong ô


– N: số ký tự cần lấy về tính từ bên trái của ô


**VD:** =LEFT(“ketoanMISA”,6) => kết quả trả về là “ketoan”, nghĩa là ta muốn lấy 6 ký tự tính từ trái qua trong “ketoanMISA”.


**>>>[Cách sử dụng hàm LEFT để trích xuất các ký tự từ chuỗi văn bản](#)**


6. Hàm excel SUBTOTAL
---------------------


Subtotal là hàm tính toán cho một nhóm con trong một danh sách hoặc bảng dữ liệu tuỳ theo phép tính mà bạn chọn lựa trong đối số thứ nhất.


Công thức: **=SUBTOTAL(9;Vùng tính tổng)**


Kế toán thường dùng hàm này khi:


– Tính tổng phát sinh trong kỳ.


– Tính tổng cho từng tài khoản cấp 1.


– Tính tổng tiền tồn cuối ngày.


Ví dụ: Trên Nhật ký chung tại cột TK Nợ/TK Có có rất nhiều tài khoản phát sinh, nếu bạn chỉ muốn xem tổng phát sinh của một tài khoản nào đó => tại dòng tổng cộng phát sinh dùng **=SUBTOTAL(9;Vùng tính tổng)**. Cuối cùng đặt lệnh Data => Filter cho vùng cần tính và lọc TK muốn xem lên, ban đã có kết quả cần tính.


[**>>>Cách sử dụng linh hoạt hàm SUBTOTAL trong Excel**](#)


Trên đây là những hàm excel thông dụng nhất, quen thuộc, dễ dùng nhất mà kế toán nào cũng nên nắm được. Để chuyên nghiệp hơn, bạn cũng nên thường xuyên trau dồi nghiệp vụ, học hỏi những công nghệ mới áp dụng trong công việc. Chúc bạn thành công!



moreExcel là phần mềm không còn xa lạ gì với dân kế toán. Nhưng không…

