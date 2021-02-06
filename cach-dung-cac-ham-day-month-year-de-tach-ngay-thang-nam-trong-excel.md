Cách dùng các hàm DAY, MONTH, YEAR để tách ngày, tháng, năm trong Excel
=======================================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-dung-cac-ham-day-month-year-de-tach-ngay-thang-nam-trong-excel/)Để tách các giá trị Ngày, Tháng, Năm trong một ngày cụ thể nào đó thì bạn nên dùng 3 hàm xử lý là hàm DAY, MONTH và YEAR. Nhóm hàm này đơn giản, dễ dùng và thường được sử dụng kết hợp với các hàm khác như SUMIFS trong việc tính doanh số, công …
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Để tách các giá trị Ngày, Tháng, Năm trong một ngày cụ thể nào đó thì bạn nên dùng 3 hàm xử lý là hàm DAY, MONTH và YEAR. Nhóm hàm này đơn giản, dễ dùng và thường được sử dụng kết hợp với các hàm khác như SUMIFS trong việc tính doanh số, công nợ theo nhiều điều kiện (trong đó có điều kiện về thời gian), sử dụng với hàm IF, sử dụng với hàm VLOOKUP, kết hợp với một chuỗi, sử dụng trong hành chính với việc xác định ngày cuối tháng để chấm công, xây dựng bảng chấm công, bảng lương, ứng dụng trong việc tìm năm sinh khi lập 1 danh sách sinh viên… Bài viết sau đây sẽ hướng dẫn các bạn cách sử dụng các hàm DAY, MONTH, YEAR nhé.**


#### 1. Hàm DAY


Hàm DAY được sử dụng trong Excel để trả về giá trị là ngày trong ô Excel. Giá trị ngày được trả về ở dạng số nguyên từ 1 đến 31.


Cú pháp hàm: **=DAY(serial\_number)**


*Trong đó:* **serial\_number**: đối số bắt buộc, là giá trị ngày, tháng, năm mà bạn muốn tách ra ngày


Ví dụ ta có bảng thời gian như hình dưới.


![](https://hddtvn.com/wp-content/uploads/2021/01/9OmSHWX.png)


*Lưu ý:* Excel lưu trữ ngày tháng ở dạng số serial liên tiếp. Theo mặc định, ngày 01/01/1900 là 1 và ngày 17/09/2012  

là 41169 bởi nó là 41169 ngày sau ngày 01/01/1900.


Theo đó:




* Tại ô A4, 41169 tương ứng với ngày 17/09/2012

* Tại ô A7, 42857 tương ứng với ngày 02/05/2017

* Tại ô A9, 43274 tương ứng với ngày 23/06/2018

* Tại ô A10, 43800 tương ứng với ngày 01/12/2019



Để tách ra được ngày của cột thời gian, áp dụng cấu trúc hàm bên trên ta có công thức cho ô B2 như sau:


**=DAY(A2)**


Sao chép công thức cho các ô còn lại của cột Ngày ta thu được kết quả:


![](https://hddtvn.com/wp-content/uploads/2021/01/eT6ODyK.png)


#### 2. Hàm MONTH


Hàm MONTH được sử dụng trong Excel để trả về giá trị là tháng trong ô Excel. Giá trị ngày được trả về ở dạng số nguyên từ 1 đến 12.


Cú pháp hàm: **=MONTH(serial\_number)**


*Trong đó:* **serial\_number**: đối số bắt buộc, là giá trị ngày, tháng, năm mà bạn muốn tách ra tháng


Vẫn với ví dụ trên, để tách ra được tháng của cột thời gian ta có công thức cho ô C2 như sau:


**=MONTH(A2)**


Sao chép công thức cho các ô còn lại của cột Tháng ta thu được kết quả:


![](https://hddtvn.com/wp-content/uploads/2021/01/4la3z9G.png)


#### 2. Hàm YEAR


Hàm YEAR được sử dụng trong Excel để trả về giá trị là năm trong ô Excel. Giá trị ngày được trả về ở dạng số nguyên từ 1900 đến 9999


Cú pháp hàm: **=YEAR(serial\_number)**


*Trong đó:* **serial\_number**: đối số bắt buộc, là giá trị ngày, tháng, năm mà bạn muốn tách ra năm


Vẫn với ví dụ trên, để tách ra được năm của cột thời gian ta có công thức cho ô D2 như sau:


**=YEAR(A2)**


Sao chép công thức cho các ô còn lại của cột Năm ta thu được kết quả:


![](https://hddtvn.com/wp-content/uploads/2021/01/4BZcjso.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm để tách ngày, tháng, năm trong Excel. Chúc các bạn thành công!*


moreĐể tách các giá trị Ngày, Tháng, Năm trong một ngày cụ thể nào đó…

