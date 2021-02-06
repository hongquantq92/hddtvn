Cách dễ nhất để phân biệt hai hàm VLOOKUP và HLOOKUP
====================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-de-nhat-de-phan-biet-hai-ham-vlookup-va-hlookup/)VLOOKUP và HLOOKUP là 2 hàm dò tìm, tham chiếu thường sử dụng trong Excel. Nhưng do cấu trúc và cách viết gần giống nhau nên rất nhiều người hay bị nhầm lẫn không phân biệt được 2 hàm này. Trong bài viết này chúng ta hãy cùng tìm hiểu cách phân biệt hàm VLOOKUP …
-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**VLOOKUP và HLOOKUP là 2 hàm dò tìm, tham chiếu thường sử dụng trong Excel. Nhưng do cấu trúc và cách viết gần giống nhau nên rất nhiều người hay bị nhầm lẫn không phân biệt được 2 hàm này. Trong bài viết này chúng ta hãy cùng tìm hiểu cách phân biệt hàm VLOOKUP với HLOOKUP và khi nào nên sử dụng những hàm này.**


Sự khác nhau giữa hai hàm tìm kiếm Vlookup và Hlookup


Để phân biệt hai hàm tìm kiếm Vlookup và Hlookup, trước hết các bạn cần biết được sự khác nhau cơ bản của hai hàm này. Khi giá trị tìm kiếm nằm trên cột dọc thì sử dụng hàm Vlookup, ngược lại nếu giá trị tìm kiếm nằm trên hàng ngang thì dùng hàm Hlookup.


Mặt khác, do cách viết của hai hàm này cũng tương tự như nhau, chỉ khác chữ cái đầu tiên (V và H) nên rất nhiều người bị nhầm lẫn. Vì thế, các bạn cần phân biệt chúng qua cách sử dụng cụ thể.


### Phân biệt cú pháp sử dụng


Cú pháp sử dụng của hai hàm tìm kiếm này như sau:




* =Vlookup(Lookup\_value,Table\_array,Col\_index\_num,[range\_lookup])

* = Hlookup(Lookup\_Value, Table\_array, Row\_index\_num, [range\_lookup])



Trong đó:




* Lookup\_value là giá trị tìm kiếm

* Table\_array là bảng giới hạn để tìm kiếm (khi thực hành các bạn cần nhấn phím F4 để cố định bảng thì

* sau này mới coppy được công thức xuống dưới)

* Col\_index\_num là số thứ tự của cột lấy giá trị tìm kiếm trong bảng

* Row\_index\_num là số thứ tự của hàng lấy giá trị tìm kiếm trong bảng

* Range\_lookup là kiểu giá trị tương đối hay tuyệt đối (TRUE và FALSE), thường lấy giá trị FALSE



Chúng ta có thể thấy trong cú pháp sử dụng của hai hàm này có nhiều đối tượng giống nhau, tuy nhiên cơ chế hoạt động của chúng thì hoàn toàn khác nhau. Cụ thể:


a. Về giá trị tìm kiếm Lookup\_value:




* Vlookup: Tìm trong cột đầu tiên của bảng

* Hlookup: Tìm trong hàng đầu tiên của bảng



b. Về kết quả cần tìm:




* Vlookup: Dóng sang phải xem ở cột nào

* Hlookup: Dóng xuống dưới xem ở hàng thứ mấy



c. Về từ khóa trong hàm:




* Vlookup: chữ V = Vertical (chiều dọc) trả kết quả theo cột

* Hlookup: chữ H = Horizontal (chiều ngang) trả kết quả theo hàng



### Cách sử dụng hai hàm tìm kiếm Vlookup và Hlookup


Để hiểu rõ và phân biệt được cách sử dụng của hai hàm tìm kiếm này, các bạn hãy theo dõi ví dụ sau đây. Kế toán có thể ứng dụng để tìm kiếm nhanh doanh thu của từng tháng hoặc chi phí cũng như lợi nhuận để làm báo cáo tổng kết năm.


Cho bảng dữ liệu như hình vẽ:


![](https://hddtvn.com/wp-content/uploads/2021/01/1-5.png)


Với ví dụ này các bạn hoàn toàn có thể sử dụng cả hai hàm Vlookup và Hlookup đều ra kết quả. Vậy với cùng một yêu cầu cách sử dụng hai hàm này khác nhau như thế nào?


#### a. Dùng hàm Vlookup để tìm kiếm doanh thu tháng 6


Với hàm Vlookup thì giá trị tìm kiếm lookup\_value phải nằm ở cột đầu tiên của bảng dữ liệu. Ta thấy “Doanh thu” nằm ở cột A, vị trí A3. Do đó nếu dùng Vlookup thì sẽ tính từ cột A trở đi. Kết quả cần tính là doanh thu của tháng 6, mà tháng 6 nằm ở vị trí cột thứ 2 tính từ cột A.


Vì thế, công thức tìm kiếm với hàm Vlookup là:


=VLOOKUP(A3,A2:H5,2,FALSE) sẽ cho ra kết quả như sau:


![](https://hddtvn.com/wp-content/uploads/2021/01/2-5.png)


#### b. Dùng hàm Hlookup để tìm kiếm doanh thu tháng 6


Với yêu cầu là tìm doanh thu tháng 6, ta thấy doanh thu tháng 6 nằm ở dòng đầu tiên của bảng dữ liệu. Do đó nếu coi lookup\_value là “Tháng 6” thì chúng ta có thể sử dụng hàm Hlookup để tìm kiếm. Kết quả tìm kiếm là doanh thu nằm ở vị trí dòng thứ 2.


Các bạn sử dụng hàm Hlookup như sau:


=HLOOKUP(B2,A2:H5,2,FALSE) được kết quả:


![](https://hddtvn.com/wp-content/uploads/2021/01/3.png)


Như vậy, cả hai hàm Vlookup và Hlookup đều cho ra kết quả giống nhau, nhưng cách sử dụng lại hoàn toàn khác. Các bạn cần phân biệt rõ lookup\_value yêu cầu tìm kiếm nội dung gì, nằm ở vị trí nào trong bảng dữ liệu. Một khi đã xác định được rõ 2 yếu tố này thì có thể xác định tiếp vị trí lấy kết quả theo hàng ngang hay theo cột dọc. Từ đó có thể lựa chọn hàm để sử dụng một cách chính xác


#### c. Tìm chi phí tháng 8 với hàm tìm kiếm Vlookup


Theo như dữ liệu cho trong bảng thì “Chi phí” nằm ở cột A, vị trí A4. Tính từ cột A thì tháng 8 sẽ nằm ở cột thứ 4. Vậy công thức tìm chi phí tháng 8 với hàm tìm kiếm Vlookup sẽ là:


=VLOOKUP(A4,A2:H5,4,FALSE)


Sau khi nhập công thức này các bạn sẽ thấy kết quả trả về như sau:


![](https://hddtvn.com/wp-content/uploads/2021/01/4-5.png)


#### d. Tìm chi phí tháng 8 với hàm tìm kiếm Hlookup


Để dùng hàm Hlookup tìm chi phí tháng 8 thì giá trị tìm kiếm Lookup\_value lúc này cần lấy không phải là “Chi phí” ở vị trí A4 nữa mà phải là “Tháng 8” ở vị trí D2. Tính theo hàng từ tháng 8 xuống dưới thì “Chi phí” nằm ở hàng thứ 3. Vậy công thức tính như sau:


=HLOOKUP(D2,A2:H5,3,FALSE)


Sau khi nhập công thức này các bạn sẽ ra được kết quả:


![](https://hddtvn.com/wp-content/uploads/2021/01/5-1.png)


Như vậy, bài viết trên đây đã hướng dẫn cách phân biệt hai hàm Vlookup và Hlookup qua 2 ví dụ minh họa cụ thể. Các bạn có thể áp dụng vào các trường hợp tương tự để sử dụng hai hàm này cho thành thạo.


Chúc các bạn sẽ phân biệt và sử dụng thành công hai hàm Vlookup và Hlookup trong Excel!


moreVLOOKUP và HLOOKUP là 2 hàm dò tìm, tham chiếu thường sử dụng trong Excel….

