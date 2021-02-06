Cách sử dụng hàm DAYS360 để tìm số ngày giữa hai ngày dựa trên một năm
======================================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-su-dung-ham-days360-de-tim-so-ngay-giua-hai-ngay-dua-tren-mot-nam/)Hàm DAYS360 được sử dụng để tính số ngày giữa hai ngày dựa trên năm có 360 ngày. Hãy theo dõi bài viết sau để biết cách sử dụng chi tiết hàm DAYS360 trong Excel. 1. Cấu trúc hàm DAYS360 Cú pháp hàm: DAYS360(start\_date,end\_date,[method]) Trong đó: Start\_date: đối số bắt buộc, là ngày bắt đầu …
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Hàm DAYS360 được sử dụng để tính số ngày giữa hai ngày dựa trên năm có 360 ngày. Hãy theo dõi bài viết sau để biết cách sử dụng chi tiết hàm DAYS360 trong Excel.**


### 1. Cấu trúc hàm DAYS360


Cú pháp hàm: **DAYS360(start\_date,end\_date,[method])**


*Trong đó:*




* **Start\_date**: đối số bắt buộc, là ngày bắt đầu để tính số ngày giữa 2 ngày đó.

* **End\_date**: đối số bắt buộc, là ngày kết thúc để tính số ngày giữa hai ngày đó.

* **Method**: đối số tùy chọn, là giá trị logic xác định sẽ dùng phương pháp của Hoa Kỳ hay của châu Âu trong tính toán.



*Lưu ý:*




* Nếu **start\_date** đến sau **end\_date**, hàm DAYS360 trả về số âm.

* Ngày nên được nhập bằng cách dùng hàm DATE hoặc nhập như là kết quả của những công thức hay hàm khác.

* Nếu đối số **Method** là FALSE hoặc bỏ qua tức là sử dụng phương pháp Hoa Kỳ (NASD). Nếu ngày bắt đầu là ngày cuối cùng của tháng, nó sẽ bằng ngày thứ 30 của tháng đó. Nếu ngày kết thúc là ngày cuối cùng của tháng và ngày bắt đầu sớm hơn ngày thứ 30 của tháng, thì ngày kết thúc sẽ bằng ngày đầu tiên của tháng tiếp theo; nếu không ngày kết thúc sẽ bằng ngày thứ 30 của tháng đó.

* Nếu đối số **Method** là TRUE tức là sử dụng phương pháp châu Âu. Ngày bắt đầu và ngày kết thúc diễn ra vào ngày 31 của tháng sẽ bằng ngày 30 của tháng đó.



### 2. Cách sử dụng hàm DAYS360


Ví dụ ta cần tính số ngày tính từ ngày 01/01/2019 đến ngày 31/01/2019 dựa trên năm có 360 ngày thì ta chỉ cần sử dụng công thức:


**=DAYS360(“01/01/2019″;”31/01/2019”)**


Kết quả sẽ trả về là 30 ngày do ta đang dựa trên giả định 1 năm có 360 ngày nên 1 tháng sẽ có 30 ngày.


![](https://hddtvn.com/wp-content/uploads/2021/01/flGWrV8.png)


Tương tự như cách dùng trong ảnh trên ta có thể thấy:




* Số ngày tính từ ngày 31/01/2019 đến ngày 01/02/2019 dựa trên năm 360 ngày là 1 ngày.

* Số ngày tính từ ngày 01/01/2019 đến ngày 31/01/2019 dựa trên năm 360 ngày là 30 ngày.

* Số ngày tính từ ngày 01/01/2019 đến ngày 31/12/2019 dựa trên năm 360 ngày là 360 ngày.

* Số ngày tính từ ngày 01/02/2019 đến ngày 01/01/2019 dựa trên năm 360 ngày là -30 ngày.



*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm DAYS360 trong Excel. Hy vọng bài viết sẽ hữu ích với các bạn trong quá trình làm việc. Chúc các bạn thành công!*


moreHàm DAYS360 được sử dụng để tính số ngày giữa hai ngày dựa trên năm…

