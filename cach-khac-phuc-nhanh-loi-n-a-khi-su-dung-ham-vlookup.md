Cách khắc phục nhanh lỗi #N/A khi sử dụng hàm VLOOKUP
=====================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-khac-phuc-nhanh-loi-n-a-khi-su-dung-ham-vlookup/)Hàm VLOOKUP có chức năng tìm kiếm giá trị trong Excel. Tuy nhiên nếu không tìm thấy kết quả thì hàm sẽ trả về giá trị lỗi #N/A. Để tránh gặp lỗi này, chúng ta có thể kết hợp hàm VLOOKUP với hàm IFERROR. Hãy theo dõi bài viết sau để biết cách thực hiện …
-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Hàm VLOOKUP có chức năng tìm kiếm giá trị trong Excel. Tuy nhiên nếu không tìm thấy kết quả thì hàm sẽ trả về giá trị lỗi #N/A. Để tránh gặp lỗi này, chúng ta có thể kết hợp hàm VLOOKUP với hàm IFERROR. Hãy theo dõi bài viết sau để biết cách thực hiện nhé.**


### 1. Cấu trúc hàm VLOOKUP


Cú pháp hàm: **=VLOOKUP(Lookup\_value, Table\_array, Col\_index\_ num, Range\_lookup)**


*Trong đó:*




* **Lookup\_value**: Giá trị cần dò tìm.

* **Table\_array**: Bảng giới hạn để dò tìm, bạn cần F4 để Fix cố định giá trị cho mục đích copy công thức tự động.

* **Col\_index\_num**: Số thứ tự của cột lấy dữ liệu trong bảng cần dò tìm.

* **Range\_lookup**: Là giá trị Logic (TRUE=1, FALSE=0) quyết định so chính xác hay so tương đổi với bảng giới hạn.



### 2. Cấu trúc hàm IFERROR


Cú pháp hàm:**=IFERROR(value, value\_if\_error)**


*Trong đó:*




* **Value**: đối số bắt buộc, là đối số cần kiểm tra xem có lỗi không.

* **Value\_if\_error**: đối số bắt buộc, là giá trị trả về nếu công thức của đối số **Value** bị lỗi. Các kiểu lỗi sau đây được đánh giá: #N/A, #VALUE!, #REF!, #DIV/0!, #NUM!, #NAME? hoặc #NULL!.



*Lưu ý:* Hàm IFERROR trong Excel nhận 2 giá trị (hoặc biểu thức) và kiểm tra, nếu:




* Nếu giá trị đầu tiên được cung cấp không đánh giá lỗi, kết quả trả về sẽ là giá trị này.

* Nếu giá trị đầu tiên được cung cấp đánh giá lỗi, kết quả trả về là giá trị thứ 2 được cung cấp.



### 3. Cách kết hợp hàm VLOOKUP và hàm IFERROR để tránh gặp lỗi #N/A


Ví dụ ta sử dụng hàm VLOOKUP trong bảng như hình dưới. Do đối tượng Đất không có trong bảng phụ nên kết quả hàm trả về lỗi #N/A.


![](https://hddtvn.com/wp-content/uploads/2021/01/wtRiqB2.png)


Để tránh gặp lỗi này, ta có thể lồng thêm hàm IFERROR vào như sau:


**=IFERROR(VLOOKUP(D10;$C$14:$D$16;2;FALSE);”Không có”)**


Chỉ cần như vậy là nếu không có kết quả tìm kiếm thì hàm sẽ trả về từ **Không có** do ta đặt.


[![](https://hddtvn.com/wp-content/uploads/2021/01/zi2sg7K.png)](https://hddtvn.com/wp-content/uploads/2021/01/zi2sg7K.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách tránh lỗi #N/A khi sử dụng hàm VLOOKUP. Hy vọng bài viết sẽ hữu ích với các bạn trong quá trình làm việc. Chúc các bạn thành công!*


moreHàm VLOOKUP có chức năng tìm kiếm giá trị trong Excel. Tuy nhiên nếu không…

