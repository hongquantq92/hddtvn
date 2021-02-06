Những nguyên nhân/cách khắc phục lỗi #N/A trong bảng tính Excel
===============================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/nhung-nguyen-nhan-cach-khac-phuc-loi-n-a-trong-bang-tinh-excel/)Lỗi #N/A xuất hiện khi công thức bạn sử dụng có chứa nội dung không có sẵn trong vùng dữ liệu, dẫn tới không thể tính toán, hoàn thành công thức được. Lỗi này thường xảy ra khi sử dụng các hàm dò tìm, tham chiếu. Đối tượng cần dò tìm, tham chiếu không có …
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Lỗi #N/A xuất hiện khi công thức bạn sử dụng có chứa nội dung không có sẵn trong vùng dữ liệu, dẫn tới không thể tính toán, hoàn thành công thức được. Lỗi này thường xảy ra khi sử dụng các hàm dò tìm, tham chiếu. Đối tượng cần dò tìm, tham chiếu không có sẵn trong vùng cần tra cứu nên báo lỗi #N/A. Về bản chất lỗi #N/A không phải là lỗi sai hàm, sai công thức mà chỉ là không tìm thấy đối tượng cần tìm. Trong bài viết này, hddtvn sẽ cùng các bạn đi tìm hiểu những nguyên nhân gây ra lỗi #N/A trong Excel và cách khắc phục nhé.**


### 1. Lỗi #N/A là gì?


#N/A là viết tắt của Not Available trong tiếng anh có nghĩa là không tồn tại, không có sẵn. Do đó lỗi này được hiểu là trong công thức bạn sử dụng có chứa nội dung không có sẵn trong vùng dữ liệu, dẫn tới không thể tính toán, hoàn thành công thức được. Lỗi này thường xảy ra khi sử dụng các hàm dò tìm, tham chiếu. Đối tượng cần dò tìm, tham chiếu không có sẵn trong vùng cần tra cứu nên báo lỗi #N/A


### 2. Khắc phục lỗi #N/A do không tìm thấy kết quả cần tìm


Ví dụ ta sử dụng hàm VLOOKUP trong bảng như hình dưới. Do đối tượng Đất không có trong bảng phụ nên kết quả hàm trả về lỗi #N/A.


![](https://hddtvn.com/wp-content/uploads/2021/01/wtRiqB2.png)


Để khắc phục, các bạn có thể thêm đối tượng Đất vào bảng kết quả tìm kiếm. Ví dụ ta sẽ thêm đối tượng Đất có thuế suất là 0%. Sau đó các bạn tiến hành sửa vùng tham chiếu trong hàm như sau:


**=VLOOKUP(D10;$C$14:$D$17;2;FALSE)**


Chỉ cần như vậy là hàm sẽ không trả về giá trị lỗi #N/A nữa mà trả về mức thuế suất là 0%.


![](https://hddtvn.com/wp-content/uploads/2021/01/c0CGfKx.png)


Hoặc nếu không muốn thêm đối tượng vào bảng tìm kiếm thì các bạn có thể lồng thêm hàm IFERROR vào để tránh gặp lỗi này như sau:


**=IFERROR(VLOOKUP(D10;$C$14:$D$16;2;FALSE);”Không có”)**


Chỉ cần như vậy là nếu không có kết quả tìm kiếm thì hàm sẽ trả về từ **Không có** do ta đặt.


![](https://hddtvn.com/wp-content/uploads/2021/01/zi2sg7K.png)


### 3. Khắc phục lỗi #N/A do khác định dạng dữ liệu


Ví dụ trong hình dưới, ta có thể thấy ngày 20/02/2019 có trong bảng giá trị tuy nhiên kết quả của hàm vẫn trả về lỗi #N/A. Điều này là do định dạng của ô **B3** và **G2** không giống nhau. Định dạng của ô **B3** là **Date** còn định dạng của ô **G2** là **Text**. Điều này khiến hàm không thể nhận dạng được dữ liệu để tính toán.


[![](https://hddtvn.com/wp-content/uploads/2021/01/GdLFyQx.png)](https://hddtvn.com/wp-content/uploads/2021/01/GdLFyQx.png)


Để khắc phục trường hợp này thì các bạn chỉ cần chỉnh định dạng của ô **G2** thành **Date** cho giống với định dạng tại ô **B3** là hàm sẽ nhận dạng được dữ liệu và tính toán một cách bình thường.


![](https://hddtvn.com/wp-content/uploads/2021/01/kc85X4M.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách khắc phục lỗi #N/A trong Excel. Hy vọng bài viết sẽ hữu ích với các bạn trong quá trình làm việc. Chúc các bạn thành công!*


moreLỗi #N/A xuất hiện khi công thức bạn sử dụng có chứa nội dung không…

