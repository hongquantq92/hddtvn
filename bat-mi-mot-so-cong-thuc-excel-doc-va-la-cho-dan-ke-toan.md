Bật mí một số công thức excel “độc” và “lạ” cho dân kế toán
===========================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/bat-mi-mot-so-cong-thuc-excel-doc-va-la-cho-dan-ke-toan/)Kế toán thường xử lý các dữ liệu, làm việc với phần mềm excel. Đây là phần mềm tính toán quen thuộc với dân kế toán nhưng đôi khi cũng gây khó khăn cho kế toán trong quá trình làm việc. Bởi việc thực hiện sai các thuật toán, sai lệnh có thể khiến kết …
------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Kế toán thường xử lý các dữ liệu, làm việc với phần mềm excel. Đây là phần mềm tính toán quen thuộc với dân kế toán nhưng đôi khi cũng gây khó khăn cho kế toán trong quá trình làm việc. Bởi việc thực hiện sai các thuật toán, sai lệnh có thể khiến kết quả sai lệch. Vì vậy, bài viết dưới đây muốn chia sẻ với các bạn một số công thức excel “độc” và “lạ”, giúp công việc của kế toán dễ dàng hơn.


1. Cộng số có chứa text
-----------------------


Thông thường với những dạng câu có số kèm với chữ khi biểu thị trên excel sẽ được phần hiểu đó là dạng text. Vì vậy chúng ta không thể thực hiện các phép tính toán trên đó được.


Với trường hợp đó, thường thì chúng ta sẽ phải dùng một số công cụ, hàm để tách phần chữ và phần số riêng. Công đoạn này có vẻ khá phức tạp nếu một bản dữ liệu chỉ toàn số kèm chữ.


Nhưng với công thức excel dưới đây, bạn có thể thực hiện tính cộng số có chữa cả text. Nếu so với phân tích trên thì có vẻ không hợp lý lắm, nhưng hãy cùng thử xem nhé:


![bảng châm công](https://hddtvn.com/wp-content/uploads/2021/01/Untitled-77.png)


Ngoài ra, chúng ta còn có thể cộng trực tiếp những số có chứa chữ giống nhau. Một lần nữa, công thức mảng được phát huy tác dụng của mình.
Tại ô M6 (hình trên) tính tổng tăng ca (C) gõ vào công thức mảng sau:


**=SUM(1*IF(RIGHT($B6:$J6)=”C”,LEFT($B6:$J6,LEN($B6:$J6)-1),0))**


* Lưu ý: $B6:$J6 là vùng chọn, “C” là ký tự cần tính tổng => Nhấn tổ hợp phím **Ctr + Shift + Enter** để xem kết quả.


Tương tự, tính tổng tăng ca (T), tại ô L7 gõ vào công thức sau:


**=SUM(1*IF(RIGHT($B6:$J6)=”T”,LEFT($B6:$J6,LEN($B6:$J6)-1),0))**


Nghĩa là ta thay “C” bằng “T” =>  sau đó nhấn tổ hợp phím **Ctr + Shift + Enter** để cho ra kết quả.


Còn khi tính tổng công thường (không chứa text) thì chúng ta vẫn dùng hàm sum để tính như bình thường, xong copy các công thức xuống các ô bên dưới.


2. Cộng trong một khoản cho trước
---------------------------------


Trong công việc đôi khi chúng ta cần tính tổng với một khoản nào đó, chẳng hạn tính tổng các giá trị nhóm lương nhân viên 3.000.000 đến 15.000.000 đồng, hay tính giá trị doanh số của nhóm có doanh thu trong tháng từ 150.000.000 đến 200.000.000 đồng. Công việc này khá quen thuộc với kế toán lương.
Ví dụ: Trong bảng lương sau, chúng ta sẽ tính tổng lương của những người làm việc ở bộ phận quản lý có mức lương từ 10.000.000 đồng đến 20.000.000 đồng, tại ô F21 chúng ta gõ vào công thức sau:


![Cộng trong một khoản cho trước](https://hddtvn.com/wp-content/uploads/2021/01/cong-thuc-1.jpg)


Công thức:


**=SUM(IF(F8:F14>=10000000,IF(F8:F14<=200000000,F8:F14,0),0))**


Đừng quên nhấn tổ hợp phím **Ctr + Shift + Enter** để cho ra kết quả. Kết quả chính xác là 37.905.385 như hình trên.


Chú ý, khi kết thức lệnh bạn không được nhấn Enter mà phải nhấn **Ctr + Shift + Enter** vì đây là công thức mảng .


3. Tính tổng theo tháng
-----------------------


Ví dụ: Tính tổng giá trị hàng bán ra trong tháng, chúng ta có bảng số liệu sau:


![tính tổng theo tháng](https://hddtvn.com/wp-content/uploads/2021/01/BAN-HANG.jpg)


Đối với yêu cầu tính tổng giá trị hàng bán ra trong tháng với bảng số liệu trên có nhiều cách tính toán nhưng bài này muốn giới thiệu cho bạn công thức excel tính nhanh, dễ nhớ nhất.


Tại ô cần tính tổng theo tháng, gõ vào công thức:


**=SUMPRODUCT((MONTH($A$4:$A$23)=1)*($F$4:$F$23))**


Trong đó MONTH($A$4:$A$23) là vùng chứa các tháng cần xử lý; $F$4:$F$23 là vùng cần tính tổng và “1” là tháng cần tính tổng.


Nếu muốn tính tổng giá trị bán hàng của tháng 2, chỉ việc thay số “1” thành số “2” cho công thức trên, các tháng còn lại tương tự. Lưu ý thêm, công thức trên chỉ đúng khi bạn cần tính toán với dữ liệu trong một năm.


4. Tính tổng đến ngày hiện tại
------------------------------


![tính tổng đến ngày hiện tại](https://hddtvn.com/wp-content/uploads/2021/01/tong-den-ngay-hien-tai2.jpg)


Khi bạn muốn tính tổng đến ngày hiện tại, tại ô cần tính toán, bạn gõ vào công thức:


**=SUMIF(A1:A9,”<=”&TODAY(),E1:E9)**


Trong đó: A1:A9 là vùng chứa tháng cần tính, E1:E9 là vùng cần tính tổng.


Do ngày hiện tại khi tính toán là ngày 6/19/2017 nên công thức trên chỉ tính tổng từ E5:E7 (tức cộng từ ngày 1/1/2017-5/26/2017) và kết quả là 24.000.


Công thức này được sử dụng chủ yếu khi bạn có một bảng kế hoạch trong tháng/ năm mà hàng hằng ngày khi có kết quả thực tế, bạn edit dữ liệu kế hoạch bằng dữ liệu thực tế và bạn muốn cộng số thực tế.


Công việc kế toán không chỉ yêu cầu chuyên môn giỏi và nghiệp vụ kế toán cũng phải thành thạo. Làm việc với các phần mềm kế toán là công việc không thể thiếu với dân kế toán. Vì vậy khi nắm được những công thức excel quen thuộc, bạn sẽ có vũ khí để xử lý công việc một cách nhanh nhất. Chúc bạn thành công!




moreKế toán thường xử lý các dữ liệu, làm việc với phần mềm excel. Đây…

