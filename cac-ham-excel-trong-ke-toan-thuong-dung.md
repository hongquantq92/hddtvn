Các hàm Excel trong kế toán thường dùng
=======================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cac-ham-excel-trong-ke-toan-thuong-dung/)Các hàm Excel thông dụng trong kế toán thường dùng lên sổ sách, tính lương, kho, nhập xuất tồn, công nợ: Hướng dẫn cách sử dụng các hàm trong Excel như: Vlookup, sumif, subtotal… 1. Hàm LEFT Cú pháp: LEFT(text,số ký tự cần lấy) VD: LEFT(“ketoanthienung”,6)=”ketoan” Nghĩa là: Tôi muốn lấy 6 ký tự trong …
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------



Các hàm Excel thông dụng trong kế toán thường dùng lên sổ sách, tính lương, kho, nhập xuất tồn, công nợ: Hướng dẫn cách sử dụng các hàm trong Excel như: Vlookup, sumif, subtotal…
------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


**1. Hàm LEFT**  

Cú pháp: LEFT(text,số ký tự cần lấy)


**VD:** LEFT(“ketoanthienung”,

6)=”ketoan”  

Nghĩa là: Tôi muốn lấy 6 ký tự trong chữ “ketoanthienung”
**VD**: LEFT(“ketoanthienung”,14)=”ketoanthienung”

  

Nghĩa là: Tôi muốn lấy 14 ký tự trong chữ “ketoanthienung”
**VD** trên Excel: LEFT(B3,2)



 ————————————————-  

  
**2. Hàm VLOOKUP**  

Cú pháp: **=VLOOKUP(giá trị dò, bảng dò, cột giá trị trả về, kiểu dò)**


– Hàm Vlookup là hàm trả về giá trị dò tìm theo cột đưa từ bảng tham chiếu lên bảng cơ sở dữ liệu theo đúng giá trị dò tìm. X=

0 là dò tìm một cách chính xác. X=1 là dò tìm một cách tương đối.
**Ví dụ như dùng để:**  

– Tìm Mã hàng hoá, tên hàng hoá từ Danh mục hàng hoá về Bảng Nhập Xuất Tồn.  

– Tìm đơn giá Xuất kho từ bên Bảng Nhập Xuất Tồn về Phiếu Xuất kho.  

– Tìm Mã TK, Tên TK từ Danh mục tài khoản về bảng CĐPS, về Sổ 131, 331…  

– Tìm số Khấu hao (Phân bổ) luỹ kế từ kỳ trước, căn cứ vào Giá trị khấu hao( phân bổ) luỹ kế (của bảng , 242, 214 )




 ———————————————————–  

  
**3. Hàm LEN**  

Cú pháp: LEN(text).


Công dụng là đếm số ký tự trong chuỗi text.


**VD**: LEN(“ketoan”)=6  

       LEN(“ketoanthienung”)=14



 —————————————————————-  

  
**4. Hàm SUMIF:**  

Cú pháp: **=SUMIF(Vùng chứa điều kiện, Điều kiện, Vùng cần tính tổng).**


– Hàm này trả về giá trị tính tổng của các ô trong vùng cần tính thoả mãn một điều kiện đưa vào.


**Ví dụ như dùng để:**

  

– Tổng hợp số liệu từ sổ NKC lên Phát sinh Nợ Phát sinh Có trên Bảng cân đối số phát sinh Tài khoản  

– Tổng hợp số liệu từ PNK, PXK lên “Bảng NHập Xuất Tồn”  

– Tổng hợp số liệu từ sổ NKC lên cột PS Nợ, PS Có của “Bảng tổng hợp phải thu, phải trả khách hàng”  

– Kết chuyển các bút toán cuối kỳ.
**Ví dụ**   = SUMIF(B3:B8,”<=8″)  

– Tính tổng của các giá trị trong vùng từ B3 đến B8 với điều kiện là các giá trị nhỏ hơn hoặc bằng 8.



 ———————————————————–  

  
**5. Hàm SUBTOTAL:**  

Cú pháp: =**SUBTOTAL(function\_num,ref1,ref2,…)**


– Hàm Subtotal là hàm tính toán cho một nhóm con trong một danh sách hoặc bảng dữ liệu tuỳ theo phép tính mà bạn chọn lựa trong đối số thứ nhất.


**Ví dụ như dùng để:**

  

– Tính tổng phát sinh trong kỳ.  

– Tính tổng cho từng tài khoản cấp 1.  

– Tính tổng tiền tồn cuối ngày…  

 

 ————————————————————  

  
**6. Hàm MID**  

Cú pháp: **MID(chuỗi ký tự, vị trí ký tự bắt đầu, số ký tự cần lấy)**


**VD:** MID(“ketoanthienung”,7,8) = thienung  

       MID(“ketoanthienung”,3,4) = toan



 ——————————————————————  

  
**7. Hàm IF**  

Cú pháp: =**If(Điều kiện, Giá trị 1, Giá trị 2).**


– Hàm IF là hàm trả về giá trị 1 nếu điều kiện đúng, Hàm trả về giá trị 2 nếu điều kiện sai.


**Ví dụ:**  

= IF(B2>=4,“DUNG”,“SAI”) = DUNG.           

= IF(B2>=5,“DUNG”,“SAI”) = SAI



 ——————————————————————  

  
**8. Hàm SUM:**  

Cú pháp:   =**SUM(Number1, Number2…)**


– Hàm Sum là hàm cộng tất cả các số trong một vùng dữ liệu được chọn.  

– Các tham số: Number1, Number2… là các số cần tính tổng.



 —————————————————————-  

  
**9. Hàm MAX:**  

Cú pháp:  **=MAX(Number1, Number2…)**


– Hàm MAX là hàm trả về số lớn nhất trong dãy được nhập.



 ——————————————————————–  

  
**10. Hàm MIN:**  

Cú pháp:   **=MIN(Number1, Number2…)**


Hàm MIN là hàm trả về số nhỏ nhất trong dãy được nhập vào.



 ——————————————————————–  

  
**11. Hàm AND:**  

Cú pháp:   **=AND(đối 1, đối 2,..).**  

– Các đối số:   Logical1, Logical2… là các biểu thức điều kiện.  

– Hàm này là Phép VÀ, chỉ đúng khi tất cả các đối số có giá trị đúng. Các đối số là các hằng, biểu thức logic. Hàm trả về giá trị TRUE (1) nếu tất cả các đối số của nó là đúng, trả về giá trị FALSE (0) nếu một hay nhiều đối số của nó là sai.  

   

**Ví dụ:** =AND(D7>0,D7<5000)


Lưu ý:  

  – Các đối số phải là giá trị logic hoặc mảng hay tham chiếu có chứa giá trị logic.  

  – Nếu đối số tham chiếu là giá trị text hoặc Null (rỗng) thì những giá trị đó bị bỏ qua.  

  – Nếu vùng tham chiếu không chứa giá trị logic thì hàm trả về lỗi #VALUE!



 ——————————————————————–  

  
**12. Hàm OR:**  

Cú pháp:   OR(đối 1, đối 2,..).  

– Các đối số: Logical1, Logical2… là các biểu thức điều kiện.  

– Hàm này là Phép HOẶC, chỉ sai khi tất cả các đối số có giá trị sai. Hàm trả về giá trị TRUE (1) nếu bất cứ một đối số nào của nó là đúng, trả về giá trị FALSE (0) nếu tất cả các đối số của nó là sai.


**– Ví dụ:**  =OR(F7>03/02/74,F7>01/01/2013)  

 



——————————————————————————

Để các bạn có thể hiểu rõ hơn và chi tiết hơn khi sử dụng các hàm này vào công việc kế toán trên Excel, mời các bạn xem thêm: **[Cách làm sổ sách kế toán trên Excel](# "cách làm sổ sách kế toán trên excel")**

**\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_**
![các hàm excel thông dụng trong kế toán](https://hddtvn.com/wp-content/uploads/2021/01/cac-ham-excel-thong-dung-trong-ke-toan.png "các hàm excel thông dụng trong kế toán")


moreCác hàm Excel trong kế toán thường dùng lên sổ sách, tính lương, quản lý kho, công nợ: Hướng dẫn cách sử dụng các hàm trong Excel như: Vlookup, sumif,…