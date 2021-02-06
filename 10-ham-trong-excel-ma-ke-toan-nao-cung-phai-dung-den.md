10 hàm trong Excel mà kế toán nào cũng phải dùng đến
====================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/10-ham-trong-excel-ma-ke-toan-nao-cung-phai-dung-den/)Kế toán, đặc biệt là kế toán đang sử dụng phần mềm excel trong việc quản lý tài chính nên biết 10 hàm excel này. Những hàm này phục vụ mật thiết đến kế toán giúp kế toán thực hiện tốt nghiệp vụ quản lý. 1. Hàm SUM, AVERAGE. • Cú pháp: = SUM(số1, số2,…, …
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Kế toán, đặc biệt là kế toán đang sử dụng phần mềm excel trong việc quản lý tài chính nên biết 10 hàm excel này. Những hàm này phục vụ mật thiết đến kế toán giúp kế toán thực hiện tốt nghiệp vụ quản lý.


1. Hàm SUM, AVERAGE.
--------------------


• Cú pháp: = SUM(số1, số2,…, số hoặc vùng dữ liệu)


Ý nghĩa: Là hàm tính tổng các giá trị.


• Cú pháp: = AVERAGE (giá trị 1, giá trị 2, giá trị 3,… giá trị n hoặc vùng dữ liệu).


Ý nghĩa: Là hàm trả về trung bình cộng các tham số đưa vào hoặc của 1 vùng dữ liệu.


2. Hàm tìm kiếm (VLOOKUP).
--------------------------


• V là viết tắt của Vertical có nghĩa là hàng dọc (thẳng đứng).


Lookup là hàm tham chiếu, tìm kiếm


VLOOKUP nghĩa là tìm kiếm theo chiều dọc, theo cột.


• Cú pháp: = VLOOKUP(x, vùng tham chiếu, cột thứ n,0).


Ý nghĩa: lấy một giá trị(x) đem so sánh theo cột của vùng tham chiếu để trả về giá trị ở cột tương ứng trong vùng tham chiếu (cột thứ n) ,


0: so sánh tương tối, 1: so sánh tuyệt tối.


Chú ý:


• Vùng tham chiếu: cột đầu tiên của tham chiếu phải bao phủ vừa đủ và toàn bộ các giá trị dò tìm. Luôn luôn phải để ở giá trị tuyệt đối.


• Cột cần lấy đếm xem nó là cột thứ bao nhiêu trong bảng tham chiếu. Khi đếm phải đếm từ trái qua phải.


3. Hàm IF
---------


• Cú pháp: = If (Logical\_test,[value\_if\_true],[value\_if\_false]) nghĩa là If( Điều kiện , giá trị 1, giá trị 2).


Ý nghĩa: Là hàm trả về giá trị 1 nếu điều kiện đúng, hàm trả về giá trị 2 nếu điều kiện sai.


![phim-tat-vang-excel](https://hddtvn.com/wp-content/uploads/2021/01/phim-tat-vang-excel.gif)


4. Hàm SUMIF
------------


• Cú pháp: = SUMIF((range, criteria,sum\_range) nghĩa là Sumif (Vùng chứa điều kiện, điều kiện, vùng cần tính tổng).


Ý nghĩa: Hàm này trả về giá trị tính tổng cảu các ô trong vùng cần tính thỏa mãn một điều kiện đưa vào.


Chú ý:


• Do tính toán trong ô của Excel, nên hàm SUMIF tính tổng này trên các phiên bản Excel 2016, Excel 2013, hay các phiên bản đời trước như Excel 2010, 2007, 2003 đều áp dụng cấu trúc hàm như nhau.


5. Hàm AND và OR
----------------


• Cú pháp: = AND((Logical1; [Logical2]; [Logical3];…) nghĩa là And(đối 1, đối 2,..).


Ý nghĩa: Hàm này là Phép VÀ, chỉ đúng khi tất cả các đối số có giá trị đúng. Các đối số là các hằng, các biểu thức logic.


Chú ý:


• Hàm AND có tối đa 256 đối số phải là các giá trị logic hay các mảng hoặc tham chiếu chứa các giá trị logic. Tất cả các giá trị sẽ bị bỏ qua nếu một đối số mảng hoặc tham chiếu có chứa văn bản hoặc ô rỗng.


• Các đối số phải là giá trị logic hoặc mảng hay tham chiếu có chứa giá trị logic.


• Nếu đối số tham chiếu là giá trị text hoặc Null (rỗng) thì những giá trị đó bị bỏ qua.


• Nếu vùng tham chiếu không chứa giá trị logic thì hàm trả về lỗi #VALUE!


• Cú pháp: = OR((Logical1; [Logical2]; [Logical3];…) nghĩa là Or(đối 1, đối 2,..).


Ý nghĩa: Hàm này là Phép HOẶC, chỉ sai khi tất cả các đối số có giá trị sai.


Ví dụ: =OR(F7>03/02/76,F7>01/01/2016).


6. Hàm COUNTIF
--------------


• Cú pháp: = COUNTIF (range,criteria).


Trong đó:


Range: là dãy dữ liệu mà bạn muốn đếm có điều kiện.


Criteria: là điều kiện để một ô được đếm.


Ý nghĩa: Hàm COUNTIF TRONG excel dùng để đếm số ô thỏa mãn điều kiện (Criteria) trong phạm vi (Range).


![phan-mem-excel](https://hddtvn.com/wp-content/uploads/2021/01/phan-mem-excel.jpg)


7. Hàm MIN, MAX
---------------


• Cú pháp: = MAX(number 1, number 2, …) .


Ý nghĩa: Trả về giá trị lớn nhất của number1, number 2,.. hoặc giá trị lớn nhất của cả 1 vùng dữ liệu số.


• Cú pháp: = MIN(number 1, number 2, …) .


Ý nghĩa: Trả về giá trị nhỏ nhất của number1, number 2,.. hoặc giá trị nhỏ nhất của cả 1 vùng dữ liệu số.


8. Hàm LEFT, RIGHT
------------------


• Cú pháp: = LEFT(chuỗi, ký tự muốn lấy).


Ý nghĩa: Tách lấy những ký tự bên trái chuỗi ký tự.


Ví dụ: = LEFT(“THANH HUE”) => Kết quả: = THANH.


• Cú pháp: = RIGHT(chuỗi, ký tự muốn lấy).


Ý nghĩa: Tách lấy những ký tự bên phải chuỗi ký tự.


Ví dụ: = LEFT(“THANH HUE”) => Kết quả: = HUE.


9. Hàm SUBTOTAL
---------------


• Cú pháp: = SUBTOTAL ((function\_num,ref1,[ref2],…).


Trong đó:


• Function\_num: Bắt buộc. Số 1-11 hay 101-111 chỉ định hàm sử dụng cho tổng phụ. 1-11 bao gồm những hàng ẩn bằng cách thủ công, còn 101-111 loại trừ chúng ra; những ô được lọc ra sẽ luôn được loại trừ.


• Ref1 Bắt buộc. Phạm vi hoặc tham chiếu được đặt tên đầu tiên mà bạn muốn tính tổng phụ cho nó.


• Ref2,… Tùy chọn. Phạm vi hoặc chuỗi được đặt tên từ 2 đến 254 mà bạn muốn tính tổng phụ cho nó.


Chú ý:


• Nếu có hàm SUBTOTAL khác lồng đặt tại các đối số ref1, ref2,.. thì các hàm lồng này sẽ bị bỏ qua không được tính nhằm tránh trường hợp tính toán 2 lần.


• Đối số Function\_num nếu từ 1 đến 11 thì hàm SUBTOTAL tính toán bao gồm cả các giá trị trong tập số liệu (hàng ẩn). Đối số Function\_num nếu từ 101 đến 111 thì hàm SUBTOTAL chỉ tính toán cho các giá trị không ẩn trong tập số liệu (bỏ qua các giá trị ẩn 0 ).


• Hàm SUBTOTAL sẽ bỏ qua không tính toán tất cả các hàng bị ẩn bởi lệnh Filter (Auto Filter) không phụ thuộc vào đối số Function\_num được dùng.


• Hàm SUBTOTAL được thiết kế để tính toán cho các cột số liệu theo chiều dọc, nó không được thiết kế để tính theo chiều ngang.


• Hàm này chỉ tính toán cho dữ liệu 2-D do vậy nếu dữ liệu tham chiếu dạng 3-D thì hàm SUBTOTAL báo lỗi #VALUE.


10. Hàm NOW
-----------


• Cú pháp: = NOW().


Cú pháp hàm NOW không có đối số nào.


Ý nghĩa: Hàm =NOW() để hiển thị ngày giờ của hệ thống trong tính.



moreKế toán, đặc biệt là kế toán đang sử dụng phần mềm excel trong việc…

