Dân kế toán luôn phải thuộc lòng những hàm Excel cơ bản này
===========================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/dan-ke-toan-luon-phai-thuoc-long-nhung-ham-excel-co-ban-nay/)Mục đích của việc tổng hợp các hàm Excel là để mọi người có thể tham khảo, và lưu lại sử dụng khi cần. Nhờ những hàm Excel này, các bạn kế toán có thể tạo cơ sở dữ liệu, bảng biểu, tham chiếu, lên các sổ sách… và lập báo cáo tài chính. Các …
-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Mục đích của việc tổng hợp các hàm Excel là để mọi người có thể tham khảo, và lưu lại sử dụng khi cần. Nhờ những hàm Excel này, các bạn kế toán có thể tạo cơ sở dữ liệu, bảng biểu, tham chiếu, lên các sổ sách… và lập báo cáo tài chính.**


Các hàm cơ bản được liệt kê cụ thể dưới đây:


#### **Hàm SUMIF**


![Dân kế toán cần nắm vững những hàm Excel nào để thành công?](https://hddtvn.com/wp-content/uploads/2021/01/cach-dung-ham-sumif-trong-excel-1.jpg "Dân kế toán cần nắm vững những hàm Excel nào để thành công?")


– Là hàm tính tổng các ô được chỉ định bởi những tiêu chuẩn đưa vào.


– Cú pháp: SUMIF (Range, Criteria, Sum\_range) nghĩa là Sumif. Đó là vùng chứa điều kiện, điều kiện, vùng cần tính tổng.


– Range: Là dãy mà bạn muốn xác định.


– Criteria: Các tiêu chuẩn mà bạn muốn tính tổng. Tiêu chuẩn này có thể là số, biểu thức hoặc chuỗi.


– Sum\_range: Là các ô thực sự cần tính tổng. Hàm này trả về giá trị tính tổng của các ô trong vùng cần tính thỏa mãn 1 điều kiện đưa vào.


Ví dụ: = SUMIF(B3:B8,”<=10’’) – Tính tổng của các giá trị trong vùng. Giá trị từ B2 đến B5 với điều kiện là các giá trị nhỏ hơn hoặc bằng 10.


**>>> [Cách sử dụng hàm SUMIF trong Excel](#)**


[**>>> Hướng dẫn dùng hàm SUMIFS để tính tổng theo nhiều điều kiện trong Excel**](#)


#### **Hàm VLOOKUP**


![Dân kế toán cần nắm vững những hàm Excel nào để thành công?](https://hddtvn.com/wp-content/uploads/2021/01/ham-vlookup-trong-excel-ke-toan-kho.png "Dân kế toán cần nắm vững những hàm Excel nào để thành công?")


– Là hàm trả về giá trị dò tìm theo cột đưa từ bảng tham chiếu lên bảng cơ sở dữ liệu theo đúng là trị dò tìm. X=0 là dò tìm 1 cách chính xác nhất. X=1 là dò tìm một cách tương đối.


– Cú pháp: Vlookup(lookup\_value, table\_array, col\_index\_num,[range\_lookup]). Nghĩa là Vlookup (Giá trị dò tìm, Bảng tham chiếu, Cột cần lấy, X).


– Các tham số: – Lookup Value: Giá trị cần đem ra so sánh để tìm kiếm.


– Table array: Bảng chứa thông tin mà dữ liệu trong bảng là dữ liệu so sánh. Vùng dữ liệu này phải tham chiếu tuyệt đối.


– Nếu như giá trị Range lookup là TRUE hoặc được bỏ qua. Thì các giá trị trong cột dùng để so sánh phải được sắp xếp tăng dần.


– Col idx num: Là số chỉ cột dữ liệu mà bạn muốn lấy trong phép so sánh.


– Range lookup: Là 1 giá trị luận lý để chỉnh cho hàm VLOOKUP tìm giá trị chính xác. Hoặc tìm giá trị gần đúng. Nếu như Range lookup là TRUE hoặc bỏ qua thì giá trị gần đúng sẽ được trả về.


***Chú ý:***


– Nếu như giá trị Lookup value nhỏ hơn giá trị nhỏ nhất trong cột đầu tiên của bảng Table array. Nó sẽ thông báo lỗi #NA.


– Ví dụ: =VLOOKUP(F11,$C$20:$D$22,2,0).


– Tìm một giá trị bằng giá trị ở ô F11 trong cột thứ nhất và lấy giá trị tương ứng ở cột 2.


[**>>>Hướng dẫn dùng hàm VLOOKUP trong Excel**](#)


[**>>>Cách sử dụng hàm HLOOKUP trong Excel**](#)


[**>>>Cách dễ nhất để phân biệt hai hàm Vlookup và Hlookup**](#)


#### **Hàm SUBTOTAL**


– Là hàm tính cho một nhóm con trong một danh sách hoặc bảng dữ liệu. Tùy theo phép tính mà bạn chọn lựa trong đối số thứ nhất.


– Cú pháp: SUBTOTAL (function\_num,ref1,ref2,…).


– Các tham số: – Function\_num là các con số từ 1 đến 11. Hay có thêm 101 đến 111 trong phiên bản Excel 2003, 2007. Quy định hàm nào sẽ được dùng để tính toán trong subtotal – Ref1, ref2,… Là các vùng địa chỉ tham chiếu mà bạn muốn thực hiện phép tính đó.


[**>>>Cách sử dụng linh hoạt hàm SUBTOTAL trong Excel**](#)


#### **Hàm AND**


– Cú pháp: AND(Logical1, Logical2,…) nghĩa là And(đối 1, đối 2,…).


– Các đối số: Logical1, Logical2,… chính là các biểu thức điều kiện.


– Hàm này là Phép VÀ, chỉ đúng khi tất cả các đối số có giá trị đúng. Các đối số là các hằng, biểu thức logic. Hàm trả về giá trị TRUE (1) nếu như tất cả các đối số của nó đúng. Và hàm trả về giá trị FALSE (0) nếu một hay nhiều đối số của nó là sai.


***Lưu ý:***


– Các đối số phải là giá trị logic hoặc mảng hay tham chiếu có chứa giá trị logic.


– Nếu đối số tham chiếu là giá trị text hoặc Null (rỗng) thì những giá trị đó bị bỏ qua.


– Nếu vùng tham chiếu không chứa giá trị logic thì hàm trả về lỗi #VALUE!


– Ví dụ như: =AND(D7>0,D7<5000).


[**>>>Mẹo dùng hàm AND và kết hợp với các hàm logic khác trong Excel**](#)


#### **Hàm MAX**


– Trả về số lớn nhất trong dãy được nhập.


– Cú pháp: MAX(Number1, Number2,…).


– Các tham số: Number1, Number2,… là dãy mà bạn muốn tìm giá trị lớn nhất ở trong đó.


[**>>> Cách tìm giá trị lớn nhất và nhỏ nhất bằng hàm MAX và MIN trong Excel**](#)


Trên đây là những thông tin về các hàm Excel dùng trong kế toán. Nếu như bạn muốn học Excel giỏi hơn, thành thạo hơn trong công việc thì bạn tham khảo nhé.


#### 


moreMục đích của việc tổng hợp các hàm Excel là để mọi người có thể…

