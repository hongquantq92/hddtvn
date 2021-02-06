Cách sử dụng hàm GCD để tìm ước số chung lớn nhất
=================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-su-dung-ham-gcd-de-tim-uoc-so-chung-lon-nhat/)Hàm GCD là một trong những hàm toán học được sử dụng phổ biến trong Excel. Trong bài viết này, hddtvn sẽ giới thiệu đến bạn đọc cách sử dụng hàm GCD để tìm ước số chung lớn nhất trong Excel. 1. Cấu trúc hàm GCD Cú pháp hàm: =GCD(number1, [number2], …) Trong đó: Number1: …
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Hàm GCD là một trong những hàm toán học được sử dụng phổ biến trong Excel. Trong bài viết này, hddtvn sẽ giới thiệu đến bạn đọc cách sử dụng hàm GCD để tìm ước số chung lớn nhất trong Excel.**


### 1. Cấu trúc hàm GCD


Cú pháp hàm: **=GCD(number1, [number2], …)**


*Trong đó:*




* **Number1**: đối số bắt buộc, là số thứ nhất cần tìm ước số chung.

* **Number2,…**: đối số tùy chọn, là các số tiếp theo cần tìm ước số chung.



*Lưu ý:*




* Hàm GCD hỗ trợ tìm ước số chung của tối đa 255 số.

* Nếu bất kỳ đối số nào không phải là số nguyên thì nó bị cắt đi phần thập phân.

* Nếu bất kỳ đối số nào không phải là số, hàm GCD trả về giá trị lỗi #VALUE!.

* Nếu bất kỳ đối số nào nhỏ hơn 0, hàm GCD trả về giá trị lỗi #NUM!.

* Tất cả các số đều chia hết cho một.

* Số nguyên tố chỉ có ước số là một và chính nó.

* Nếu một tham số của GCD >=2^53, hàm GCD trả về giá trị lỗi #NUM!.



### 2. Cách sử dụng hàm GCD


Ví dụ nếu ta cần tìm ước số chung lớn nhất của 21 và 35. Áp dụng cấu trúc hàm như trên ta có công thức như sau:


**=GCD(21;35)**


![](https://hddtvn.com/wp-content/uploads/2021/01/msl3pSQ.png)


Tương tự để tìm ước số chung lớn nhất của 16, 40 và 64 ta có công thức như sau:


**=GCD(16;40;64)**


![](https://hddtvn.com/wp-content/uploads/2021/01/cBTmTKK.png)


Trường hợp tìm ước số chung lớn nhất của 21,54 và 35,12. Hàm sẽ cắt đi phần thập phần để trở thành tìm ước số chung lớn nhất của 21 và 35.


Trường hợp tìm ước số chung lớn nhất của 21 và -35. Hàm sẽ trả về giá trị lỗi #NUM! vì có đối số âm.


![](https://hddtvn.com/wp-content/uploads/2021/01/kO12i0S.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn sử dụng hàm GCD để tìm ước chung lớn nhất. Hy vọng bài viết sẽ hữu ích với các bạn. Chúc các bạn thành công!*


moreHàm GCD là một trong những hàm toán học được sử dụng phổ biến trong…

