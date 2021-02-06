Mẹo tạo thông báo hợp đồng sắp hết hạn trong Excel
==================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/meo-tao-thong-bao-hop-dong-sap-het-han-trong-excel/)Quản lý thời hạn hợp đồng lao động là một công việc phức tạp tại những đơn vị có quy mô lớn. Trong bài viết này, hddtvn xin chia sẻ với các bạn mẹo tạo thông báo hợp đồng sắp hết hạn trong Excel nhằm hỗ trợ những bạn hành chính nhân sự dễ dàng …
-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Quản lý thời hạn hợp đồng lao động là một công việc phức tạp tại những đơn vị có quy mô lớn. Trong bài viết này, hddtvn xin chia sẻ với các bạn mẹo tạo thông báo hợp đồng sắp hết hạn trong Excel nhằm hỗ trợ những bạn hành chính nhân sự dễ dàng theo dõi tình hình hợp đồng của người lao động.**


Ví dụ ta có danh sách lao động kèm ngày hết hạn hợp đồng trong hình dưới. Để tạo thông báo hợp đồng sắp hết hạn các bạn hãy làm theo các bước sau.


**Bước 1:** Đầu tiên các bạn cần bôi đen toàn bộ cột Ngày hết hạn hợp đồng. Sau đó chọn thể **Home** trên thanh công cụ. Tiếp theo các bạn chọn **Conditional Formatting** tại mục **Styles**. Thanh cuộn hiện ra các bạn chọn **New Rule**.


![](https://hddtvn.com/wp-content/uploads/2021/01/6ssfLzZ.png)


**Bước 2:** Lúc này, cửa sổ **New Formatting Rule** hiện ra. Tại mục **Select a Rule Type** các bạn chọn **Use a formula to determine which cells to format**. Sau đó tại mục **Format values where this formula is true** các bạn nhập công thức sau:


**=DATEDIF(TODAY(),D2,”D”)<30**


Với công thức trên, thời gian thông báo sẽ là 30 ngày trước khi hợp động hết hạn.


*Trong đó:*




* Hàm **DATEDIF** đếm số ngày giữa hai mốc thời gian.

* **D2** là ô chứa giá trị Ngày hết hạn trên bảng tính.

* **30** là số ngày thông báo theo mong muốn (có thể thay tùy ý theo nhu cầu thực tế).



![](https://hddtvn.com/wp-content/uploads/2021/01/sOXzUNz.png)


#### **Bước 3:** Sau khi đã nhập xong công thức, các bạn chọn mục **Format** để chọn màu thông báo.


![](https://hddtvn.com/wp-content/uploads/2021/01/mWqFIgx.png)


Lúc này, cửa sổ Format Cells xuất hiện. Các bạn chọn thẻ **Fill** sau đó chọn màu thông báo tại mục **Background Color**. Cuối cùng nhấn **OK** để hoàn tất.


![](https://hddtvn.com/wp-content/uploads/2021/01/WfqCF5z.png)


Kết quả ta thu được là có 3 nhân viên có hợp đồng sắp hết hạn trong vòng 30 ngày tới đã được highlight trong hình dưới.


[![](https://hddtvn.com/wp-content/uploads/2021/01/gBBTQZM.png)](https://hddtvn.com/wp-content/uploads/2021/01/gBBTQZM.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách tạo thông báo hợp đồng sắp hết hạn trong Excel. Hy vọng bài viết sẽ hữu ích với các bạn trong quá trình làm việc. Chúc các bạn thành công!*


moreQuản lý thời hạn hợp đồng lao động là một công việc phức tạp tại…

