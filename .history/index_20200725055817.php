<?php
session_start();

if (!isset($_SESSION['login'])) {
  header("Location: login.php");
  exit;
}

require 'functions.php';
$mahasiswa = query("SELECT * FROM mahasiswa");

// ketika tombol cari diklik
if (isset($_POST['cari'])) {
  $mahasiswa = cari($_POST['keyword']);
}

?>
<!DOCTYPE html>
<html lang="en">

<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Daftar Mahasiswa</title>
  <style>
    .loader {
      width: 100px;
      position: absolute;
      top: 118px;
      left: 210px;
      z-index: -1;
      display: none;
    }

    @media print {

      .logout,
      .tambah,
      .form-cari,
      .aksi {
        display: none;
      }
    }
  </style>
</head>

<body>
  <a href="logout.php">Logout</a> | <a href="cetak.php" target="_blank">Cetak</a>
  <h3>Daftar Mahasiswa</h3>

  <a href="tambah.php" class="tambah">Tambah data mahasiswa</a>
  <br><br>

  <form action="" method="POST" class="form-cari">
    <input type="text" name="keyword" size="40" placeholder="masukkan keyword pencarian.." autocomplete="off" autofocus class="keyword" id="keyword">
    <button type="submit" name="cari" class="tombol-cari">Cari!</button>
  </form>
  <br>

  <div class="container">
    <table border="1" cellpadding="10" cellspacing="0">
      <tr>
        <th>#</th>
        <th>Gambar</th>
        <th class="aksi">Aksi</th>
        <!-- <th>NRP</th> -->
        <th>Nama</th>
        <!-- <th>Email</th> -->
        <!-- <th>Jurusan</th> -->
      </tr>

      <?php if (empty($mahasiswa)) : ?>
        <tr>
          <td colspan="4">
            <p style="color: red; font-style: italic;">data mahasiswa tidak ditemukan!</p>
          </td>
        </tr>
      <?php endif; ?>

      <?php $i = 1;
      foreach ($mahasiswa as $m) : ?>
        <tr>
          <td><?= $i++; ?></td>
          <td><img src="img/<?= $m['gambar']; ?>" width="60"></td>
          <td><?= $m['nama']; ?></td>
          <td>
            <a href="detail.php?id=<?= $m['id']; ?>">lihat detail</a>
          </td>
        </tr>
      <?php endforeach; ?>
    </table>
  </div>

  <script src="js/script.js"></script>
</body>

</html>