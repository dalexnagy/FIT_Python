-- phpMyAdmin SQL Dump
-- version 4.9.5deb2ubuntu0.1~esm1
-- https://www.phpmyadmin.net/
--
-- Host: localhost:3306
-- Generation Time: Sep 18, 2023 at 06:05 AM
-- Server version: 8.0.34-0ubuntu0.20.04.1
-- PHP Version: 7.4.3-4ubuntu2.19

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
SET AUTOCOMMIT = 0;
START TRANSACTION;
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8mb4 */;

--
-- Database: `CyclingData`
--
CREATE DATABASE IF NOT EXISTS `CyclingData` DEFAULT CHARACTER SET utf8mb4 COLLATE utf8mb4_0900_ai_ci;
USE `CyclingData`;

-- --------------------------------------------------------

--
-- Table structure for table `RideStats`
--

CREATE TABLE `RideStats` (
  `FileName` varchar(64) NOT NULL COMMENT 'FIT File Name',
  `StartTime` datetime NOT NULL,
  `EndTime` datetime NOT NULL,
  `RideTimeSecs` int NOT NULL,
  `TotalTimeSecs` int NOT NULL,
  `TotalDistMeters` float(10,2) NOT NULL,
  `SpeedMaxMetersSec` float(10,2) NOT NULL,
  `SpeedAvgMetersSec` float(10,2) NOT NULL,
  `AltitudeMaxMeters` float(10,2) NOT NULL,
  `AltitudeMinMeters` float(10,2) NOT NULL,
  `AscentMeters` int NOT NULL,
  `AscentGradePct` float(10,2) NOT NULL,
  `DescentMeters` int NOT NULL,
  `DescentGradePct` float(10,2) NOT NULL,
  `FrontGearChanges` int NOT NULL,
  `RearGearChanges` int NOT NULL,
  `FinalElemntCharge` int NOT NULL,
  `FinalDI2Charge` int NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Indexes for dumped tables
--

--
-- Indexes for table `RideStats`
--
ALTER TABLE `RideStats`
  ADD UNIQUE KEY `FileName` (`FileName`),
  ADD KEY `FileName_2` (`FileName`);
COMMIT;

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
