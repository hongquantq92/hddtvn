Cách dùng hàm TRUNC để loại bỏ phần thập phân và rút ngắn giá trị số
====================================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-dung-ham-trunc-de-loai-bo-phan-thap-phan-va-rut-ngan-gia-tri-so/)Hàm TRUNC được sử dụng rộng rãi trong Excel nhằm mục đích loại bỏ phần thập phân và rút ngắn giá trị số. Điểm khác giữa hàm TRUNC với các hàm Excel khác là hàm này không làm tròn các giá trị. Ví dụ, ô của bạn có giá trị là 4.68. Để loại bỏ …
-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Hàm TRUNC được sử dụng rộng rãi trong Excel nhằm mục đích loại bỏ phần thập phân và rút ngắn giá trị số. Điểm khác giữa hàm TRUNC với các hàm Excel khác là hàm này không làm tròn các giá trị. Ví dụ, ô của bạn có giá trị là 4.68. Để loại bỏ các chữ số ở phần thập phân, chúng ta sử dụng hàm TRUNC và kết quả trả về sẽ là 4. Để biết thêm nhiều công dụng khác của hàm TRUNC, mời** **bạn đọc bài viết sau nhé.**


![Cách dùng hàm TRUNC để loại bỏ phần thập phân và rút ngắn giá trị số](https://hddtvn.com/wp-content/uploads/2021/01/trunc.png)


### 1. Cấu trúc hàm TRUNC


Cú pháp hàm: **=TRUNC(number, [num\_digits])**


*Trong đó:*




* **Number**: đối số bắt buộc, là số cần làm tròn.

* **Num\_digits**: đối số tùy chọn, là một số xác định độ chính xác của việc cắt bớt. Giá trị mặc định của **num\_digits** là 0.



*Lưu ý:* Hàm TRUNC và INT có điểm chung là cả hai hàm đều trả về số nguyên. TRUNC loại bỏ phần phân số của một số. INT làm tròn xuống số nguyên gần nhất dựa trên giá trị của phần phân số của số đó. INT và TRUNC chỉ khác nhau khi sử dụng số âm: TRUNC(-5.3) trả về -5, nhưng INT(-5.3) trả về -6 vì -6 là số nhỏ hơn.


### 2. Cách sử dụng hàm TRUNC


#### a. Sử dụng hàm TRUNC để loại bỏ phần thập phần


Ví dụ ta có các số thập phần như hình dưới. Để loại bỏ phần thập phân đi, các bạn chỉ cần sử dụng công thức


**=TRUNC(A2)**


Sao chép công thức cho các số bên dưới ta sẽ thu được kết quả là phần thập phân của các số đã bị cắt đi. Hàm TRUNC sẽ chỉ đơn giản là cắt đi phần thập phân chứ sẽ không làm tròn gì cả.


![](https://hddtvn.com/wp-content/uploads/2021/01/IvuNQzo.png)


#### b. Sử dụng hàm TRUNC để cắt bớt phần thập phân


Cũng với ví dụ trên, nếu bạn cần cắt phần thập phân của tất cả các số thành chỉ còn 2 chữ số thì ta sẽ sử dụng công thức sau:


=TRUNC(A2;2)


Sao chép công thức cho các số còn lại ta sẽ thu được kết quả là phần thập phân của các số đã bị cắt chỉ còn 2 chữ số.


![](https://hddtvn.com/wp-content/uploads/2021/01/wknTEfC.png)


#### c. Sử dụng hàm TRUNC để rút ngắn số


Trường hợp này sẽ khá hiếm gặp. Khi bạn để đối số **Num\_digits** là số âm, hàm sẽ cắt các số ở bên trái dấu thập phân. Tuy nhiên, nó không thay đổi số lượng chữ số, mà thay thế chúng bằng số không.


![](https://hddtvn.com/wp-content/uploads/2021/01/rgDKa3r.png)


#### d. Dùng hàm TRUNC để xóa giờ trong thời gian


Trường hợp trong ô thời gian của bạn có chứ cả số giờ. Để loại bỏ phần giờ này đi và chỉ để lại phần ngày tháng, các bạn sử dụng hàm TRUNC bằng công thức sau:


**=TRUNC(A2)**


Sao chép công thức cho các ô còn lại ta sẽ thu được kết quả là phần giờ đã bị loại bỏ. Ô kết quả sẽ chỉ còn lại phần ngày tháng.


![](https://hddtvn.com/wp-content/uploads/2021/01/hh5r8eo.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm TRUNC trong Excel. Hy vọng bài viết sẽ hữu ích với các bạn trong quá trình làm việc. Chúc các bạn thành công!*


moreHàm TRUNC được sử dụng rộng rãi trong Excel nhằm mục đích loại bỏ phần…

