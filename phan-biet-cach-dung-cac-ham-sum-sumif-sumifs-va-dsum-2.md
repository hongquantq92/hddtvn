Phân biệt cách dùng các hàm SUM, SUMIF, SUMIFS và DSUM
======================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/phan-biet-cach-dung-cac-ham-sum-sumif-sumifs-va-dsum-2/)Khi cần tính tổng, có lẽ hàm đầu tiên bạn nghĩ tới là SUM. Tuy nhiên khi nhập công thức =SUM trong một ô trong Excel, bạn sẽ nhận được rất nhiều gợi ý công thức bắt đầu bằng SUM. Bạn chắc sẽ phân vân không biết chúng khác nhau thế nào? Cách dùng cụ …
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Khi cần tính tổng, có lẽ hàm đầu tiên bạn nghĩ tới là SUM. Tuy nhiên khi nhập công thức =SUM trong một ô trong Excel, bạn sẽ nhận được rất nhiều gợi ý công thức bắt đầu bằng SUM. Bạn chắc sẽ phân vân không biết chúng khác nhau thế nào? Cách dùng cụ thể ra sao? Bài viết này sẽ giúp bạn phân biệt cách dùng 4 hàm tính tổng thông dụng nhất là SUM, SUMIF, SUMIFS và DSUM. Mỗi một hàm có chức năng riêng để tính tổng cho dữ liệu, hãy cùng tìm hiểu từng hàm một nhé.**


![Phân biệt cách dùng các hàm SUM, SUMIF, SUMIFS và DSUM](https://hddtvn.com/wp-content/uploads/2021/01/phan-biet-cac-h-am.png)


### 1. Hàm SUM


Hàm SUM có chức năng tính tổng nhiều ô riêng lẻ, tính tổng trong một phạm vi hoặc thậm chí nhiều phạm vi trong một lần.


Cú pháp hàm: **=SUM(number1,[number2],…)**


*Trong đó:*




* **number1**: đối số bắt buộc, là số đầu tiên bạn muốn thêm vào.

* **number2-255**: đối số tùy chọn, là số thứ 2 đến thứ 255 mà bạn muốn cộng.



Ví dụ để tính tổng tiền hàng trong bảng dữ liệu của hình dưới ta có công thức như sau:


**=SUM(F2:F10)**


![Phân biệt cách dùng các hàm SUM, SUMIF, SUMIFS và DSUM](https://hddtvn.com/wp-content/uploads/2021/01/22-4.png "Phân biệt cách dùng các hàm SUM, SUMIF, SUMIFS và DSUM")


### 2. Hàm SUMIF


Hàm SUMIF thực hiện tính tổng dựa trên một điều kiện.


Cú pháp hàm: **=SUMIF(range; criteria;[sum\_range])**


*Trong đó:*




* **Range**: Là vùng được chọn có chứa các ô điều kiện

* **Criteria**: Là điều kiện để thực hiện hàm này

* **Sum\_range**: Vùng cần tính tổng



Ví dụ cần tính tổng tiền hàng của các mặt hàng màu đỏ ta có công thức như sau:


**=SUMIF(B2:B10;B3;F2:F10)**


![](https://hddtvn.com/wp-content/uploads/2021/01/23-5.png)


### 3. Hàm SUMIFS


Hàm SUMIFS tính tổng dựa trên nhiều điều kiện, khác với hàm SUMIF ở trên chỉ có thể tính tổng dựa trên 1 điều kiện. Bạn cũng có thể sử dụng hàm này với chỉ một điều kiện.


Cú pháp hàm: **=SUMIFS(sum\_range, criteria\_range1, criteria1, [criteria\_range2, criteria2], …)**


*Trong đó:*




* **Sum\_range**: đối số bắt buộc, là vùng cần tính tổng.

* **Criteria\_range1**: đối sốt bắt buộc, là vùng chứa các ô điều kiện thứ 1.

* **Criteria**: đối số bắt buộc, là điều kiện thứ 1.

* **Criteria\_range2…**: đối số tùy chọn, là vùng chứa các ô điều kiện thứ 2 trở đi.

* **Criteria…**: đối số tùy chọn, là điều kiện thứ 2 trở đi.



Ví dụ để tính tổng tiền hàng của các mặt hàng màu đỏ và có đơn giá nhỏ hơn 5.000.000 ta có công thức như sau:


**=SUMIFS(F2:F10;B2:B10;B3;D2:D10;”<5000000″)**


![](https://hddtvn.com/wp-content/uploads/2021/01/24-6.png)


### 4. Hàm DSUM


Hàm DSUM là hàm tính tổng của một trường hoặc 1 cột để thỏa mãn điều kiện đưa ra.


Cú pháp hàm: **=DSUM(database;field;criteria)**


*Trong đó:*




* **Database** là cơ sở dữ liệu được tạo từ 1 phạm vi ô. Danh sách dữ liệu này sẽ chứa các dữ liệu là các trường, gồm trường để kiểm tra điều kiện và trường để tính tổng. Danh sách chứa hàng đầu tiên là tiêu đề cột.

* **Field** chỉ rõ tên cột dùng để tính tổng các số liệu. Có thể nhập tên tiêu đề cột trong dấu ngoặc kép hoặc là một số thể hiện vị trí cột trong danh sách không trong dấu ngoặc kép (ví dụ số 1 là cột đầu tiên, số 2 là cột thứ 2… trong database) hoặc tham chiếu đến tiêu đề cột mà các bạn muốn tính tổng.

* **Criteria** là phạm vi ô có chứa điều kiện mà các bạn muốn hàm DSUM kiểm tra.



Ví dụ để tính tổng tiền hàng của các mặt hàng đỏ và trắng, ta có thể sử dụng hàm DSUM thay vì 2 lần hàm SUMIF như sau:


**=DSUM(A1:F10;F1;K8:K10)**


![](https://hddtvn.com/wp-content/uploads/2021/01/25-5.png)


*Như vậy, bài viết trên đã phân biệt sự khác nhau giữa các hàm SUM, SUMIF, SUMIFS và DSUM. Hy vọng bài viết sẽ hữu ích với các bạn trong quá trình làm việc. Chúc các bạn thành công!*


#### 


moreKhi cần tính tổng, có lẽ hàm đầu tiên bạn nghĩ tới là SUM. Tuy…

