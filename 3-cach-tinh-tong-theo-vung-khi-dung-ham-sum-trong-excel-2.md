3 Cách tính tổng theo vùng khi dùng hàm SUM trong Excel
=======================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/3-cach-tinh-tong-theo-vung-khi-dung-ham-sum-trong-excel-2/)Hàm SUM là hàm cơ bản nhưng có thể ứng dụng để tính tổng trong rất nhiều trường hợp khác nhau. Trong bài viết này, hddtvn sẽ chia sẻ với các bạn tính tổng nhiều vùng, tính tổng của các vùng giao nhau và tính tổng vùng liên tục với hàm SUM nhé. 1. Cấu trúc của …
---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Hàm SUM là hàm cơ bản nhưng có thể ứng dụng để tính tổng trong rất nhiều trường hợp khác nhau. Trong bài viết này, hddtvn sẽ chia sẻ với các bạn tính tổng nhiều vùng, tính tổng của các vùng giao nhau và tính tổng vùng liên tục với hàm SUM nhé.**


![3 Cách tính tổng theo vùng khi dùng hàm SUM trong Excel](https://hddtvn.com/wp-content/uploads/2021/01/3-bien-the-ham-sum-3.png)


### 1. Cấu trúc của hàm SUM


Cú pháp hàm: **=SUM(number1,[number2],…)**


*Trong đó:*




* **number1**: đối số bắt buộc, là số đầu tiên bạn muốn thêm vào. Số đó có thể là 4, tham chiếu ô như B6, hoặc ô phạm vi như B2:B8.

* **number2-255**: đối số tùy chọn, là số thứ 2 đến thứ 255 mà bạn muốn cộng. Số đó có thể là 4, tham chiếu ô như B6, hoặc ô phạm vi như B2:B8.



Ví dụ để tính tổng từ ô F3 đến F7 ta có công thức tính như sau:


**=SUM(F3:F7)**


![3 Cách tính tổng theo vùng khi dùng hàm SUM trong Excel](https://hddtvn.com/wp-content/uploads/2021/01/27-3.png "3 Cách tính tổng theo vùng khi dùng hàm SUM trong Excel")


### 2. Sử dụng hàm SUM để tính tổng nhiều vùng


Ta có thể sử dụng hàm SUM để tính tổng của nhiều vùng cùng một lúc. Ví dụ để tính tổng từ ô B3 đến B7 với từ ô D3 đến D7 và từ F3 đến F7 thì ta sẽ có công thức tính như sau:


**=SUM(B3:B7;D3:D7;F3:F7)**


![](https://hddtvn.com/wp-content/uploads/2021/01/28-2.png)


### 3. Sử dụng hàm SUM để tính tổng của các vùng giao nhau


Ví dụ để tính tổng của vùng giao nhau giữa hai vùng B5:F5 và D3:E7, ta có công thức như sau:


**=SUM(B5:F5 D3:E7)**


Lưu ý ngắn cách giữa 2 vùng bằng dấu cách Space. Kết quả ta sẽ thu được vùng giao nhau giữa 2 vùng là ô D5 và E5 có tổng giá trị là 13 + 14 = 27.


![](https://hddtvn.com/wp-content/uploads/2021/01/30-3.png)


### 4. Sử dụng hàm SUM để tính tổng vùng liên tục


Gần tương tự với cách tính tổng 1 vùng lớn, ta có thể sử dụng toán tử (:) để phân tách ô đầu tiên khỏi ô cuối cùng khi bạn tham chiếu đến phạm vi ô liên tục trong công thức. Chẳng hạn, để tính tổng vùng từ B3 tới F7, ngoài cách nhập =SUM(B3:F7), ta có thể nhập theo những phương thức khác. Ví dụ như sau:




* **=SUM(B3:C3:E7:F7)**

* **=SUM(B3:B4:F6:F7)**

* **=SUM(B6:B7:F3:F4)**



![3 Cách tính tổng theo vùng khi dùng hàm SUM trong Excel](https://hddtvn.com/wp-content/uploads/2021/01/31-3.png "3 Cách tính tổng theo vùng khi dùng hàm SUM trong Excel")


*Như vậy, bài viết trên đã chia sẻ với các bạn 3 cách tính tổng theo vùng với hàm SUM trong Excel. Hy vọng bài viết sẽ hữu ích với các bạn trong quá trình làm việc. Chúc các bạn thành công!*


#### 


moreHàm SUM là hàm cơ bản nhưng có thể ứng dụng để tính tổng trong…

