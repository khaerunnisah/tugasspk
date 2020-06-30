-- phpMyAdmin SQL Dump
-- version 5.0.2
-- https://www.phpmyadmin.net/
--
-- Host: 127.0.0.1
-- Waktu pembuatan: 30 Jun 2020 pada 12.28
-- Versi server: 10.4.11-MariaDB
-- Versi PHP: 7.4.6

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
START TRANSACTION;
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8mb4 */;

--
-- Database: `db_saham`
--

-- --------------------------------------------------------

--
-- Struktur dari tabel `tbl_bobot`
--

CREATE TABLE `tbl_bobot` (
  `kd_bobot` varchar(15) NOT NULL DEFAULT '',
  `jns_bobot` varchar(40) DEFAULT NULL,
  `nilai` float DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- --------------------------------------------------------

--
-- Struktur dari tabel `tbl_kriteria`
--

CREATE TABLE `tbl_kriteria` (
  `kd_kriteria` varchar(15) NOT NULL DEFAULT '',
  `nm_kriteria` varchar(60) DEFAULT NULL,
  `atribut` varchar(40) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- --------------------------------------------------------

--
-- Struktur dari tabel `tbl_nilai`
--

CREATE TABLE `tbl_nilai` (
  `Id` int(11) NOT NULL,
  `kd_saham` int(11) DEFAULT NULL,
  `nm_saham` varchar(255) DEFAULT NULL,
  `aset` decimal(10,2) DEFAULT NULL,
  `laba_bersih` decimal(10,2) DEFAULT NULL,
  `laba_kotor` decimal(10,2) DEFAULT NULL,
  `laba_usaha` decimal(10,2) DEFAULT NULL,
  `pendapatan` decimal(10,2) DEFAULT NULL,
  `per` decimal(10,2) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

--
-- Dumping data untuk tabel `tbl_nilai`
--

INSERT INTO `tbl_nilai` (`Id`, `kd_saham`, `nm_saham`, `aset`, `laba_bersih`, `laba_kotor`, `laba_usaha`, `pendapatan`, `per`) VALUES
(2, 0, NULL, '1.00', '5.00', '3.00', '4.00', '3.00', '2.00');

-- --------------------------------------------------------

--
-- Struktur dari tabel `tbl_saham`
--

CREATE TABLE `tbl_saham` (
  `kd_saham` varchar(11) NOT NULL DEFAULT '',
  `nm_saham` varchar(60) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

--
-- Indexes for dumped tables
--

--
-- Indeks untuk tabel `tbl_bobot`
--
ALTER TABLE `tbl_bobot`
  ADD PRIMARY KEY (`kd_bobot`);

--
-- Indeks untuk tabel `tbl_kriteria`
--
ALTER TABLE `tbl_kriteria`
  ADD PRIMARY KEY (`kd_kriteria`);

--
-- Indeks untuk tabel `tbl_nilai`
--
ALTER TABLE `tbl_nilai`
  ADD PRIMARY KEY (`Id`);

--
-- Indeks untuk tabel `tbl_saham`
--
ALTER TABLE `tbl_saham`
  ADD PRIMARY KEY (`kd_saham`);

--
-- AUTO_INCREMENT untuk tabel yang dibuang
--

--
-- AUTO_INCREMENT untuk tabel `tbl_nilai`
--
ALTER TABLE `tbl_nilai`
  MODIFY `Id` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=3;
COMMIT;

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
