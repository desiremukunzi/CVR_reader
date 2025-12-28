-- phpMyAdmin SQL Dump
-- version 5.2.1
-- https://www.phpmyadmin.net/
--
-- Host: 127.0.0.1
-- Generation Time: Dec 28, 2025 at 01:21 AM
-- Server version: 10.4.32-MariaDB
-- PHP Version: 8.1.25

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
START TRANSACTION;
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8mb4 */;

--
-- Database: `fdap_new`
--

-- --------------------------------------------------------

--
-- Table structure for table `aircrafts`
--

CREATE TABLE `aircrafts` (
  `id` int(11) NOT NULL,
  `aircraft_type_id` int(11) NOT NULL,
  `call_sign` varchar(20) NOT NULL,
  `empty_weight` decimal(9,2) DEFAULT 8000.00,
  `max_takeoff_weight` decimal(9,2) DEFAULT 13000.00,
  `status_id` int(11) DEFAULT 1,
  `manufactured_date` date DEFAULT NULL,
  `created_at` timestamp NOT NULL DEFAULT current_timestamp(),
  `updated_at` timestamp NOT NULL DEFAULT current_timestamp() ON UPDATE current_timestamp()
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

--
-- Dumping data for table `aircrafts`
--

INSERT INTO `aircrafts` (`id`, `aircraft_type_id`, `call_sign`, `empty_weight`, `max_takeoff_weight`, `status_id`, `manufactured_date`, `created_at`, `updated_at`) VALUES
(1, 2, 'UNO-565P', 8000.00, 13000.00, 1, NULL, '2025-12-26 19:20:50', '2025-12-26 19:22:20'),
(2, 1, 'UNO-560P', 8000.00, 13000.00, 1, NULL, '2025-12-26 19:20:50', '2025-12-26 19:22:26'),
(3, 1, 'UNO-561P', 8000.00, 13000.00, 1, NULL, '2025-12-26 19:20:50', '2025-12-26 19:22:31');

-- --------------------------------------------------------

--
-- Table structure for table `aircraft_categories`
--

CREATE TABLE `aircraft_categories` (
  `id` int(11) NOT NULL,
  `name` varchar(20) NOT NULL,
  `created_at` timestamp NOT NULL DEFAULT current_timestamp(),
  `updated_at` timestamp NOT NULL DEFAULT current_timestamp() ON UPDATE current_timestamp()
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

--
-- Dumping data for table `aircraft_categories`
--

INSERT INTO `aircraft_categories` (`id`, `name`, `created_at`, `updated_at`) VALUES
(1, 'helicopter', '2025-12-26 13:05:11', '2025-12-26 13:05:11'),
(2, 'fixed_wing', '2025-12-26 13:05:11', '2025-12-26 13:05:11');

-- --------------------------------------------------------

--
-- Table structure for table `aircraft_types`
--

CREATE TABLE `aircraft_types` (
  `id` int(11) NOT NULL,
  `aircraft_category_id` int(11) NOT NULL,
  `type` varchar(20) NOT NULL,
  `created_at` timestamp NOT NULL DEFAULT current_timestamp(),
  `updated_at` timestamp NOT NULL DEFAULT current_timestamp() ON UPDATE current_timestamp()
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

--
-- Dumping data for table `aircraft_types`
--

INSERT INTO `aircraft_types` (`id`, `aircraft_category_id`, `type`, `created_at`, `updated_at`) VALUES
(1, 1, 'MI-17V-5', '2025-12-26 13:05:38', '2025-12-26 13:05:38'),
(2, 1, 'MI-17-1V', '2025-12-26 13:05:38', '2025-12-26 13:05:38');

-- --------------------------------------------------------

--
-- Table structure for table `anomalies`
--

CREATE TABLE `anomalies` (
  `id` int(11) NOT NULL,
  `flight_id` int(11) NOT NULL,
  `parameter_MI_17V_5_name` varchar(20) NOT NULL,
  `phase_of_flight_id` int(11) NOT NULL,
  `total_anomalies` int(5) NOT NULL,
  `created_at` timestamp NOT NULL DEFAULT current_timestamp(),
  `updated_at` timestamp NOT NULL DEFAULT current_timestamp() ON UPDATE current_timestamp()
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

--
-- Dumping data for table `anomalies`
--

INSERT INTO `anomalies` (`id`, `flight_id`, `parameter_MI_17V_5_name`, `phase_of_flight_id`, `total_anomalies`, `created_at`, `updated_at`) VALUES
(31, 10, 'Fcp', 2, 8, '2025-12-28 00:15:38', '2025-12-28 00:15:38'),
(32, 10, 'Xcpl', 3, 1, '2025-12-28 00:15:38', '2025-12-28 00:15:38'),
(33, 10, 'Pedals', 2, 17, '2025-12-28 00:15:38', '2025-12-28 00:15:38'),
(34, 10, 'X_lat', 2, 4, '2025-12-28 00:15:38', '2025-12-28 00:15:38'),
(35, 10, 'X_lat', 3, 3, '2025-12-28 00:15:38', '2025-12-28 00:15:38'),
(36, 11, 'Fcp', 1, 4, '2025-12-28 00:15:55', '2025-12-28 00:15:55'),
(37, 11, 'Fcp', 2, 34, '2025-12-28 00:15:55', '2025-12-28 00:15:55'),
(38, 11, 'Xcpl', 1, 6, '2025-12-28 00:15:55', '2025-12-28 00:15:55'),
(39, 11, 'Pedals', 2, 13, '2025-12-28 00:15:55', '2025-12-28 00:15:55'),
(40, 11, 'PITCH', 2, 14, '2025-12-28 00:15:55', '2025-12-28 00:15:55');

-- --------------------------------------------------------

--
-- Table structure for table `checklist_types`
--

CREATE TABLE `checklist_types` (
  `id` int(11) NOT NULL,
  `name` varchar(20) NOT NULL,
  `created_at` timestamp NOT NULL DEFAULT current_timestamp(),
  `updated_at` timestamp NOT NULL DEFAULT current_timestamp() ON UPDATE current_timestamp()
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

--
-- Dumping data for table `checklist_types`
--

INSERT INTO `checklist_types` (`id`, `name`, `created_at`, `updated_at`) VALUES
(1, 'STARTING WITH AC-GPU', '2025-12-27 19:56:56', '2025-12-27 19:56:56'),
(2, 'STARTING WITH DC-GPU', '2025-12-27 19:56:56', '2025-12-27 19:56:56'),
(3, 'STARTING WITHOUT GPU', '2025-12-27 19:56:56', '2025-12-27 19:56:56');

-- --------------------------------------------------------

--
-- Table structure for table `crews`
--

CREATE TABLE `crews` (
  `id` int(11) NOT NULL,
  `rank` varchar(11) DEFAULT NULL,
  `first_name` varchar(30) NOT NULL,
  `last_name` varchar(30) NOT NULL,
  `code` varchar(10) NOT NULL,
  `crew_type_id` int(11) NOT NULL,
  `status_id` int(2) NOT NULL,
  `created_at` timestamp NOT NULL DEFAULT current_timestamp(),
  `updated_at` timestamp NOT NULL DEFAULT current_timestamp() ON UPDATE current_timestamp()
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

--
-- Dumping data for table `crews`
--

INSERT INTO `crews` (`id`, `rank`, `first_name`, `last_name`, `code`, `crew_type_id`, `status_id`, `created_at`, `updated_at`) VALUES
(28, 'Capt', 'Miquera', 'C UMUHOZA', 'S1', 2, 1, '2025-12-26 12:54:28', '2025-12-26 12:54:28'),
(29, 'Capt', 'Moses', 'MURASHI', 'S2', 2, 1, '2025-12-26 12:54:28', '2025-12-26 12:54:28'),
(30, 'Capt', 'Aaron', 'KAMUGISHA', 'S3', 2, 1, '2025-12-26 12:54:28', '2025-12-26 12:54:28'),
(31, 'Lt', 'Josia', 'N RUGEMA', 'S4', 2, 1, '2025-12-26 12:54:28', '2025-12-26 12:54:28'),
(32, 'Lt', 'Eloge', 'N NYIRINGANGO', 'S5', 2, 1, '2025-12-26 12:54:28', '2025-12-26 12:54:28'),
(33, 'Maj', 'Rafiki', 'KARUME', 'F1', 3, 1, '2025-12-26 12:54:28', '2025-12-26 12:54:28'),
(34, 'Maj', 'Edward', 'MUTESA', 'F2', 3, 1, '2025-12-26 12:54:28', '2025-12-26 12:54:28'),
(35, 'Capt', 'Ronald', 'NSANZUMUHIRE', 'F3', 3, 1, '2025-12-26 12:54:28', '2025-12-26 12:54:28'),
(36, 'Lt', 'Edward', 'NSHUTI', 'F4', 3, 1, '2025-12-26 12:54:28', '2025-12-26 12:54:28'),
(37, 'Capt', 'JB', 'MUSIRIKARE', 'F5', 3, 1, '2025-12-26 12:54:28', '2025-12-26 12:54:28'),
(38, 'Maj', 'I', 'KARINGANIRE', 'F6', 3, 1, '2025-12-26 12:54:28', '2025-12-26 12:54:28'),
(40, 'Lt Col', 'Edgard', 'NYAKAYIRU', 'P1', 1, 0, '2025-12-26 19:14:02', '2025-12-26 19:14:02'),
(41, 'Lt Col', 'Alain', 'RUZINDANA', 'P2', 1, 1, '2025-12-26 19:14:02', '2025-12-26 19:14:02'),
(42, 'Maj', 'Serge', 'KAMANYA', 'P3', 1, 1, '2025-12-26 19:14:02', '2025-12-26 19:14:02'),
(43, 'Capt', 'Fred', 'Didas RWIGEMA', 'P4', 1, 1, '2025-12-26 19:14:02', '2025-12-26 19:14:02'),
(44, 'Lt Col', 'Achille', 'NSENGIYUMVA', 'P5', 1, 1, '2025-12-26 19:14:02', '2025-12-26 19:14:02'),
(45, 'Maj', 'Sam', 'RUNANIRA', 'P6', 1, 1, '2025-12-26 19:14:02', '2025-12-26 19:14:02'),
(46, 'Capt', 'E', 'IBAMBASI', 'P7', 1, 1, '2025-12-26 19:14:02', '2025-12-26 19:14:02');

-- --------------------------------------------------------

--
-- Table structure for table `crew_types`
--

CREATE TABLE `crew_types` (
  `id` int(11) NOT NULL,
  `name` varchar(20) NOT NULL,
  `created_at` timestamp NOT NULL DEFAULT current_timestamp(),
  `updated_at` timestamp NOT NULL DEFAULT current_timestamp() ON UPDATE current_timestamp()
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

--
-- Dumping data for table `crew_types`
--

INSERT INTO `crew_types` (`id`, `name`, `created_at`, `updated_at`) VALUES
(1, 'PIC', '2025-12-26 12:50:16', '2025-12-26 12:50:16'),
(2, 'SIC', '2025-12-26 12:50:16', '2025-12-26 12:50:16'),
(3, 'FE', '2025-12-26 12:50:16', '2025-12-26 12:50:16');

-- --------------------------------------------------------

--
-- Table structure for table `exceedances`
--

CREATE TABLE `exceedances` (
  `id` int(11) NOT NULL,
  `flight_id` int(11) NOT NULL,
  `parameter_MI_17V_5_name` varchar(20) NOT NULL,
  `number_of_exceedances` int(3) NOT NULL,
  `created_at` timestamp NOT NULL DEFAULT current_timestamp(),
  `updated_at` timestamp NOT NULL DEFAULT current_timestamp() ON UPDATE current_timestamp()
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

--
-- Dumping data for table `exceedances`
--

INSERT INTO `exceedances` (`id`, `flight_id`, `parameter_MI_17V_5_name`, `number_of_exceedances`, `created_at`, `updated_at`) VALUES
(37, 10, 'Roll', 1, '2025-12-28 00:15:38', '2025-12-28 00:15:38'),
(38, 10, 'N1/N2 Split', 2, '2025-12-28 00:15:38', '2025-12-28 00:15:38'),
(39, 10, 'iChips', 1, '2025-12-28 00:15:38', '2025-12-28 00:15:38'),
(40, 10, 'iOP1', 2, '2025-12-28 00:15:38', '2025-12-28 00:15:38');

-- --------------------------------------------------------

--
-- Table structure for table `flights`
--

CREATE TABLE `flights` (
  `id` int(11) NOT NULL,
  `flight_date` date NOT NULL,
  `aircraft_id` int(11) NOT NULL,
  `PIC` varchar(3) NOT NULL,
  `SIC` varchar(3) NOT NULL,
  `FE` varchar(3) NOT NULL,
  `sortie` int(2) NOT NULL,
  `flight_type_id` int(11) DEFAULT 1,
  `checks_not_complied` int(3) DEFAULT NULL,
  `compliance_percentage` decimal(5,2) DEFAULT NULL,
  `continuous_exceedances` int(6) NOT NULL,
  `discrete_exceedances` int(7) NOT NULL,
  `anomalies` int(7) NOT NULL,
  `anomalies_percentage` float(7,2) NOT NULL,
  `created_at` timestamp NOT NULL DEFAULT current_timestamp(),
  `updated_at` timestamp NOT NULL DEFAULT current_timestamp() ON UPDATE current_timestamp()
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

--
-- Dumping data for table `flights`
--

INSERT INTO `flights` (`id`, `flight_date`, `aircraft_id`, `PIC`, `SIC`, `FE`, `sortie`, `flight_type_id`, `checks_not_complied`, `compliance_percentage`, `continuous_exceedances`, `discrete_exceedances`, `anomalies`, `anomalies_percentage`, `created_at`, `updated_at`) VALUES
(10, '2025-10-01', 3, 'P6', 'S4', 'F6', 1, 1, 7, 95.90, 3, 3, 33, 0.49, '2025-12-28 00:15:38', '2025-12-28 00:15:38'),
(11, '2025-07-31', 3, 'P3', 'S3', 'F5', 2, 1, 8, 95.30, 0, 0, 71, 1.85, '2025-12-28 00:15:55', '2025-12-28 00:15:55');

-- --------------------------------------------------------

--
-- Table structure for table `flight_types`
--

CREATE TABLE `flight_types` (
  `id` int(11) NOT NULL,
  `name` varchar(20) NOT NULL,
  `created_at` timestamp NOT NULL DEFAULT current_timestamp(),
  `updated_at` timestamp NOT NULL DEFAULT current_timestamp() ON UPDATE current_timestamp()
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

--
-- Dumping data for table `flight_types`
--

INSERT INTO `flight_types` (`id`, `name`, `created_at`, `updated_at`) VALUES
(1, 'cargo', '2025-12-26 13:46:46', '2025-12-26 13:46:46'),
(2, 'normal_passenger', '2025-12-26 13:46:46', '2025-12-26 13:46:46'),
(3, 'VIP_passenger', '2025-12-26 13:46:46', '2025-12-26 13:46:46');

-- --------------------------------------------------------

--
-- Table structure for table `missed_checks`
--

CREATE TABLE `missed_checks` (
  `id` int(11) NOT NULL,
  `flight_id` int(11) NOT NULL,
  `checklist_type_id` int(11) NOT NULL,
  `checklist_item_position` int(4) NOT NULL,
  `created_at` timestamp NOT NULL DEFAULT current_timestamp(),
  `updated_at` timestamp NOT NULL DEFAULT current_timestamp() ON UPDATE current_timestamp()
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

--
-- Dumping data for table `missed_checks`
--

INSERT INTO `missed_checks` (`id`, `flight_id`, `checklist_type_id`, `checklist_item_position`, `created_at`, `updated_at`) VALUES
(196, 10, 1, 8, '2025-12-28 00:15:38', '2025-12-28 00:15:38'),
(197, 10, 1, 11, '2025-12-28 00:15:38', '2025-12-28 00:15:38'),
(198, 10, 1, 119, '2025-12-28 00:15:38', '2025-12-28 00:15:38'),
(199, 10, 1, 136, '2025-12-28 00:15:38', '2025-12-28 00:15:38'),
(200, 10, 1, 141, '2025-12-28 00:15:38', '2025-12-28 00:15:38'),
(201, 10, 1, 165, '2025-12-28 00:15:38', '2025-12-28 00:15:38'),
(202, 10, 1, 167, '2025-12-28 00:15:38', '2025-12-28 00:15:38'),
(203, 11, 3, 21, '2025-12-28 00:15:55', '2025-12-28 00:15:55'),
(204, 11, 3, 73, '2025-12-28 00:15:55', '2025-12-28 00:15:55'),
(205, 11, 3, 82, '2025-12-28 00:15:55', '2025-12-28 00:15:55'),
(206, 11, 3, 136, '2025-12-28 00:15:55', '2025-12-28 00:15:55'),
(207, 11, 3, 142, '2025-12-28 00:15:55', '2025-12-28 00:15:55'),
(208, 11, 3, 150, '2025-12-28 00:15:55', '2025-12-28 00:15:55'),
(209, 11, 3, 163, '2025-12-28 00:15:55', '2025-12-28 00:15:55'),
(210, 11, 3, 170, '2025-12-28 00:15:55', '2025-12-28 00:15:55');

-- --------------------------------------------------------

--
-- Table structure for table `parameters`
--

CREATE TABLE `parameters` (
  `id` int(11) NOT NULL,
  `MI_17V_5_name` varchar(20) NOT NULL,
  `MI_17_1V_name` varchar(20) DEFAULT NULL,
  `description` varchar(100) DEFAULT NULL,
  `discrete` tinyint(1) NOT NULL,
  `aircraft_type_id` int(11) NOT NULL,
  `created_at` timestamp NOT NULL DEFAULT current_timestamp(),
  `updated_at` timestamp NOT NULL DEFAULT current_timestamp() ON UPDATE current_timestamp()
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

--
-- Dumping data for table `parameters`
--

INSERT INTO `parameters` (`id`, `MI_17V_5_name`, `MI_17_1V_name`, `description`, `discrete`, `aircraft_type_id`, `created_at`, `updated_at`) VALUES
(7, 'IAS', NULL, 'Exceeded speed limits based on altitude and current aircraft weight', 0, 1, '2025-12-26 13:15:38', '2025-12-26 13:15:38'),
(8, 'Alt', NULL, 'Exceeded altitude limit (>6000m or >4800m with aircraft weight >11100kg)', 0, 1, '2025-12-26 13:15:38', '2025-12-26 13:15:38'),
(9, 'ROLL', NULL, 'Exceeded roll limits based on altitude and current aircraft weight', 0, 1, '2025-12-26 13:15:38', '2025-12-26 13:15:38'),
(10, 'PITCH', NULL, 'Outside the valid range (-20 to 20)', 0, 1, '2025-12-26 13:15:38', '2025-12-26 13:15:38'),
(11, 'Fcp', NULL, 'Outside the valid range (0-15)', 0, 1, '2025-12-26 13:15:38', '2025-12-26 13:15:38'),
(12, 'N1/N2 Split', NULL, 'Excessive split between N1 and N2 values ', 0, 1, '2025-12-26 13:15:38', '2025-12-26 13:15:38'),
(13, 'N1', NULL, 'Out of(72-80) with NR between 56-70 or out of(80-101.15) when airborne', 0, 1, '2025-12-26 13:15:38', '2025-12-26 13:15:38'),
(14, 'N2', NULL, 'Out of(72-80) with NR between 56-70 or out of(80-101.15) when airborne', 0, 1, '2025-12-26 13:15:38', '2025-12-26 13:15:38'),
(15, 'Nmr', NULL, 'Abnormal NR in Idle (BEO) or Airborne', 0, 1, '2025-12-26 13:15:38', '2025-12-26 13:15:38'),
(16, 'iAPr/p', NULL, 'Roll and pitch channels disengagement', 1, 1, '2025-12-26 13:15:38', '2025-12-26 13:15:38'),
(17, 'iChips', NULL, 'Chips in gear boxes', 1, 1, '2025-12-26 13:15:38', '2025-12-26 13:15:38'),
(18, 'iEMG1', NULL, 'LH engine emergency power', 1, 1, '2025-12-26 13:15:38', '2025-12-26 13:15:38'),
(19, 'iEMG2', NULL, 'RH engine emergency power', 1, 1, '2025-12-26 13:15:38', '2025-12-26 13:15:38'),
(20, 'iF_gen1', NULL, 'Generator 1 fault', 1, 1, '2025-12-26 13:15:38', '2025-12-26 13:15:38'),
(21, 'iF_gen2', NULL, 'Generator 2 fault', 1, 1, '2025-12-26 13:15:38', '2025-12-26 13:15:38'),
(22, 'iF_pump1', NULL, 'Fuel pump 1 activation', 1, 1, '2025-12-26 13:15:38', '2025-12-26 13:15:38'),
(23, 'iF_pump2', NULL, 'Fuel pump 2 activation', 1, 1, '2025-12-26 13:15:38', '2025-12-26 13:15:38'),
(24, 'iF_pumpS', NULL, 'Standby fuel pump activation', 1, 1, '2025-12-26 13:15:38', '2025-12-26 13:15:38'),
(25, 'iFire_KO-50', NULL, 'Fire detected in KO-50 heater', 1, 1, '2025-12-26 13:15:38', '2025-12-26 13:15:38'),
(26, 'iFire_mgb', NULL, 'Fire detected in MGB', 1, 1, '2025-12-26 13:15:38', '2025-12-26 13:15:38'),
(27, 'iFire_v1', NULL, 'Fire in Vibration Sensor 1', 1, 1, '2025-12-26 13:15:38', '2025-12-26 13:15:38'),
(28, 'iFire_v2', NULL, 'Fire in Vibration Sensor 2', 1, 1, '2025-12-26 13:15:38', '2025-12-26 13:15:38'),
(29, 'iFire1', NULL, 'Fire detected in Engine 1', 1, 1, '2025-12-26 13:15:38', '2025-12-26 13:15:38'),
(30, 'iFire2', NULL, 'Fire detected in Engine 2', 1, 1, '2025-12-26 13:15:38', '2025-12-26 13:15:38'),
(31, 'iHSaux', NULL, 'Auxiliary hydraulic system status', 1, 1, '2025-12-26 13:15:38', '2025-12-26 13:15:38'),
(32, 'iHSmain', NULL, 'Main hydraulic system status', 1, 1, '2025-12-26 13:15:38', '2025-12-26 13:15:38'),
(33, 'inFT1', NULL, 'Fuel tank 1 level warning', 1, 1, '2025-12-26 13:15:38', '2025-12-26 13:15:38'),
(34, 'inFT2', NULL, 'Fuel tank 2 level warning', 1, 1, '2025-12-26 13:15:38', '2025-12-26 13:15:38'),
(35, 'iOP_mgb', NULL, 'Oil pressure MGB warning', 1, 1, '2025-12-26 13:15:38', '2025-12-26 13:15:38'),
(36, 'iOP1', NULL, 'Oil pressure Engine 1 warning', 1, 1, '2025-12-26 13:15:38', '2025-12-26 13:15:38'),
(37, 'iOP2', NULL, 'Oil pressure Engine 2 warning', 1, 1, '2025-12-26 13:15:38', '2025-12-26 13:15:38'),
(38, 'iQTmin', NULL, 'LOW FUEL 270L reserve warning', 1, 1, '2025-12-26 13:15:38', '2025-12-26 13:15:38'),
(92, 'Xcpl', NULL, 'Throttle', 0, 1, '2025-12-27 23:44:50', '2025-12-27 23:44:50'),
(93, 'Pedals', NULL, 'Pedals', 0, 1, '2025-12-27 23:44:50', '2025-12-27 23:45:38'),
(94, 'X_lat', '', 'Lateral movement of cyclic', 0, 1, '2025-12-27 23:44:50', '2025-12-27 23:45:23'),
(95, 'X_long', NULL, 'Longitudinal movement of cyclic', 0, 1, '2025-12-27 23:44:50', '2025-12-27 23:44:50'),
(97, 'NZ', NULL, 'Vertical acceleration', 0, 1, '2025-12-27 23:44:50', '2025-12-27 23:44:50'),
(98, 'T1', NULL, 'Gas temperature for left engine', 0, 1, '2025-12-27 23:44:50', '2025-12-27 23:44:50'),
(99, 'T2', NULL, 'Gas temperature for left engine', 0, 1, '2025-12-27 23:44:50', '2025-12-27 23:44:50');

-- --------------------------------------------------------

--
-- Table structure for table `phase_of_flights`
--

CREATE TABLE `phase_of_flights` (
  `id` int(11) NOT NULL,
  `name` varchar(20) NOT NULL,
  `created_at` timestamp NOT NULL DEFAULT current_timestamp(),
  `updated_at` timestamp NOT NULL DEFAULT current_timestamp() ON UPDATE current_timestamp()
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

--
-- Dumping data for table `phase_of_flights`
--

INSERT INTO `phase_of_flights` (`id`, `name`, `created_at`, `updated_at`) VALUES
(1, 'before_takeoff', '2025-12-27 23:23:57', '2025-12-27 23:23:57'),
(2, 'airborne', '2025-12-27 23:23:57', '2025-12-27 23:23:57'),
(3, 'after_landing', '2025-12-27 23:23:57', '2025-12-27 23:23:57');

-- --------------------------------------------------------

--
-- Table structure for table `starting_without_gpu_checklist`
--

CREATE TABLE `starting_without_gpu_checklist` (
  `id` int(11) NOT NULL,
  `name` varchar(100) NOT NULL,
  `created_at` timestamp NOT NULL DEFAULT current_timestamp(),
  `updated_at` timestamp NOT NULL DEFAULT current_timestamp() ON UPDATE current_timestamp()
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

--
-- Dumping data for table `starting_without_gpu_checklist`
--

INSERT INTO `starting_without_gpu_checklist` (`id`, `name`, `created_at`, `updated_at`) VALUES
(2, 'MI-17V-5 HELICOPTER START UP WITH BATTERY & INFLIGHT CHECKLIST', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(3, 'PRE APU START', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(4, 'Instruments and all switches As required', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(5, 'Battery 1 and 2 On and Check', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(6, 'Circuit Breakers On as required', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(7, 'FDR and CVR On', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(8, 'Headsets Connected', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(9, 'Intercom Readability Check', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(10, 'Aircraft records on board and filled', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(11, 'Overhead hatch closed', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(12, 'Windscreens Clean', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(13, 'Seatbelts Fastened', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(14, 'Pneumatic System Check Pressure', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(15, 'Pedals Neutral', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(16, 'Cyclic Stick Neutral', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(17, 'Landing gear brakes applied', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(18, 'Collective pitch fully down', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(19, 'Throttle Twist Grip Fully Left', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(20, 'Friction Clutch Adjusted', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(21, 'Separate Throttle Lever Middle and Latched', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(22, 'Main Rotor Brake Fully Down', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(23, 'Engine Shut Down Levers Backward', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(24, 'Fuel Quantity as per mission', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(25, 'Warning Lights Check', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(26, 'Fire extinguishing System Check operation', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(27, 'Voice Warning System Check operation', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(28, 'Inverter On', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(29, 'EGT Indicators Test', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(30, 'Engine vibration system test', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(31, 'Inverter Off', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(32, 'Fire extinguishing System On', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(33, 'Service pump On', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(34, 'Generators Off', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(35, 'Fire Fuel shut off valve Check and  Open', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(36, 'Engine Shut Down Levers Backward', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(37, 'Pre APU Start Checklist Completed', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(38, 'APU START CHECKLIST', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(39, 'Startup clearance ATC Request', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(40, 'Ground crew Signal', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(41, 'APU Start', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(42, 'APU Parameters Check as Required', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(43, 'APU Generator On', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(44, 'Rectifiers On', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(45, 'All circuit breakers and transfer pumps On', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(46, 'Startup checklist Completed', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(47, 'ENGINE START UP', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(48, 'Start up Area Clear', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(49, 'Anti collision light  On', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(50, 'Engine selection as per wind direction', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(51, 'Ground crew  Signal', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(52, 'Engine  Start button press', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(53, 'HP cock open', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(54, 'Stopwatch Set', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(55, 'Warm up engine 1 to 2 minutes', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(56, 'Engine Parameters as required', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(57, 'Start second engine as above', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(58, 'Warm up second engine 1 to 2 minutes', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(59, 'Idle parameters Check', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(60, 'DPD On', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(61, 'EGT Air Check', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(62, 'FUNCTIONAL CHECK', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(63, 'Hydraulic system Check', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(64, 'Controls Response Check', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(65, 'EEG Test', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(66, 'Partial acceleration Check', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(67, 'ENGINE OPERATIONAL MODE', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(68, 'Throttle Twist Grip Fully Right', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(69, 'Generators Test and On', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(70, 'Inverters auto position', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(71, 'APU Off', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(72, 'Navigation Lights On', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(73, 'Blade tip lights Night and Poor visibility On', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(74, 'Cabin Lighting Night and Poor visibility On', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(75, 'Formation lights Night and Poor visibility On', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(76, 'Voice Warning System On', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(77, 'Gyro Horizons On', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(78, 'Compass System On and Check', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(79, 'ADF On', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(80, 'GPS On', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(81, 'Transponder On', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(82, 'ELT Arming', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(83, 'Global satellite tracking system On', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(84, 'TCAS On', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(85, 'Radio Altimeter On and Check', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(86, 'Baro Altimeter Set', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(87, 'Pitch Limit System Check', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(88, 'Auto Pilot On and Check', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(89, 'Main Rotor Control speed Check range', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(90, 'Main Rotor RPM Set 95 percent', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(91, 'Collective  Pitch Down', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(92, 'Engine Startup checklist Completed', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(93, 'BEFORE TAXI', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(94, 'Taxing clearance ATC Request', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(95, 'Area Clear', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(96, 'Chocks   Removed and Stowed', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(97, 'Crew And Pax Briefed', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(98, 'Doors and Windows    Closed', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(99, 'Cargo Compartment  Secured', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(100, 'Autopilot Off', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(101, 'Before Taxing Checklist  Completed', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(102, 'LINEUP POSITION', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(103, 'Area Clear', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(104, 'Obstacles in TakeOff Direction Absent', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(105, 'Gyro Same readings', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(106, 'Heading Set', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(107, 'Autopilot ON', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(108, 'Type of TakeOFF Decide', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(109, 'Request for TakeOFF Performed', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(110, 'Stopwatch Press', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(111, 'BEFORE TAKEOFF', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(112, 'Fuel Selector Service', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(113, 'Fuel Pump Lights  Check Off', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(114, 'Transponder Alt', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(115, 'Autopilot On', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(116, 'Engine And Transmission Checked', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(117, 'Before Takeoff Checklist Completed', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(118, 'AFTER TAKEOFF', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(119, 'Main rotor 95 plus or minus 2 percent', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(120, 'DPDS Above 50 meters OFF', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(121, 'Fuel consumption Monitored', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(122, 'Monitor every 15 to 20 Minutes', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(123, 'After takeoff checklist     Completed', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(124, 'IN FLIGHT CHECKS', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(125, 'Power setting As per graph', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(126, 'Fuel quantity Check', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(127, 'Other Parameters Normal operation', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(128, 'Flight Following Call', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(129, 'PRE LANDING CHECKS', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(130, 'Landing clearance Request', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(131, 'Runway Condition Known', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(132, 'Autopilot Alt channel Off', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(133, 'Cargo Secured', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(134, 'Landing lights On', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(135, 'Fuel quantity Check', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(136, 'Compass Matched', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(137, 'DPDS On At 50 meters', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(138, 'Type of Landing Decide', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(139, 'Runway Exit and Parking As instructed', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(140, 'Parameters Within Limit', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(141, 'Main rotor 95 plus or minus 2 percent', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(142, 'Landing Gear Brakes Released', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(143, 'DPDS On', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(144, 'Cargo compartment Secured', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(145, 'Crew Briefed', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(146, 'Before landing checklist Completed', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(147, 'AFTER LANDING', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(148, 'Collective FULLY DOWN', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(149, 'Auto pilot OFF', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(150, 'parking as required', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(151, 'After landing checklist COMPLETED', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(152, 'ENGINE SHUT DOWN', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(153, 'Landing gear Brake Apply', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(154, 'Consumers Off', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(155, 'Rectifiers Off', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(156, 'Inverters Neutral', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(157, 'Generators Off', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(158, 'Throttle Fully Left', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(159, 'Left and Right Fuel Pumps Off', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(160, 'Engines Cooling 1 to 2 Minutes', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(161, 'Engine Shutdown Levers Backward', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(162, 'Stopwatch Press', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(163, 'Ngg equal 0 at 35 Second Minimum', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(164, 'MAIN ROTOR STOP', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(165, 'At less or equal to 15 % Main Rotor Slowly apply Main Rotor brake', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(166, 'At Main Rotor stop  Controls move back and forth', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(167, 'Fire extinguisher switch Off', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(168, 'Fuel shut off valves Leave open position', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(169, 'Service tank pump Off', '2025-12-26 13:20:55', '2025-12-26 13:20:55'),
(170, 'Batteries one and two Off', '2025-12-26 13:20:55', '2025-12-26 13:20:55');

-- --------------------------------------------------------

--
-- Table structure for table `starting_with_ac_gpu_checklist`
--

CREATE TABLE `starting_with_ac_gpu_checklist` (
  `id` int(11) NOT NULL,
  `name` varchar(100) NOT NULL,
  `created_at` timestamp NOT NULL DEFAULT current_timestamp(),
  `updated_at` timestamp NOT NULL DEFAULT current_timestamp() ON UPDATE current_timestamp()
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

--
-- Dumping data for table `starting_with_ac_gpu_checklist`
--

INSERT INTO `starting_with_ac_gpu_checklist` (`id`, `name`, `created_at`, `updated_at`) VALUES
(2, 'MI-17V-5 HELICOPTER START UP WITH AC GPU & INFLIGHT CHECKLIST', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(3, 'PRE APU START', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(4, 'Instruments and all switches As required', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(5, 'Battery 1 and 2 Check and On', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(6, 'GPU Connected check and On', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(7, 'Rectifiers On', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(8, 'Circuit Breakers On', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(9, 'FDR and CVR On', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(10, 'Headsets Connected', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(11, 'Intercom Readability Check', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(12, 'ADF On', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(13, 'GPS On', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(14, 'Transponder On', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(15, 'ELT Arming', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(16, 'Global satellite tracking system On', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(17, 'TCAS On', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(18, 'Aircraft records on board and filled', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(19, 'Overhead hatch closed', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(20, 'Windscreens Clean', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(21, 'Seatbelts Fastened', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(22, 'Pneumatic System Check Pressure', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(23, 'Pedals Neutral', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(24, 'Cyclic Stick Neutral', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(25, 'Landing gear brakes applied', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(26, 'Collective pitch fully down', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(27, 'Throttle Twist Grip Fully Left', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(28, 'Friction Clutch Adjusted', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(29, 'Separate Throttle Lever Middle and Latched', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(30, 'Main Rotor Brake Fully Down', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(31, 'Engine Shut Down Levers Backward', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(32, 'Fuel Quantity Check', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(33, 'Warning Lights Check', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(34, 'Fire extinguishing System Check', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(35, 'Voice Warning System Check', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(36, 'EGT Indicators Test', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(37, 'Engine vibration system test', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(38, 'Fire extinguishing System On', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(39, 'All pumps On', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(40, 'Generators Off', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(41, 'Fire Fuel shut off valve Check and  Open', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(42, 'Engine Shut Down Levers Backward', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(43, 'Pre APU Start Checklist Completed', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(44, 'APU START CHECKLIST', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(45, 'Startup clearance ATC Request', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(46, 'Ground crew Signal', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(47, 'APU Start', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(48, 'APU Parameters Check as Required', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(49, 'APU Generator On', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(50, 'GPU Off and disconnected', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(51, 'Startup checklist Completed', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(52, 'ENGINE START UP', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(53, 'Start up Area Clear', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(54, 'Anti collision light  On', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(55, 'Engine selection   as per wind direction', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(56, 'Ground crew  Signal', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(57, 'Engine  Start button    press', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(58, 'HP cock open', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(59, 'Stop watch Set', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(60, 'Warm up engine 1 to 2 minutes', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(61, 'Engine Parameters as required', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(62, 'Start second engine as above', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(63, 'Warm up second engine 1 to 2 minutes', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(64, 'Idle parameters Check', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(65, 'DPD On', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(66, 'EGT Air Check', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(67, 'FUNCTIONAL CHECK', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(68, 'Hydraulic system Check', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(69, 'Controls Response Check', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(70, 'EEG Test', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(71, 'Partial acceleration Check', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(72, 'ENGINE OPERATIONAL MODE', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(73, 'Throttle Twist Grip Fully Right', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(74, 'Generators Test and On', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(75, 'APU Generator Off', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(76, 'Inverters Auto position', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(77, 'APU Off', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(78, 'Navigation Lights On', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(79, 'Blade tip lights Night and Poor visibility On', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(80, 'Cabin Lighting Night and Poor visibility On', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(81, 'Formation lights Night and Poor visibility On', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(82, 'Voice Warning System On', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(83, 'Gyro Horizons On', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(84, 'Compass System On and Check', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(85, 'Radio Altimeter On and Check', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(86, 'Baro Altimeter Set', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(87, 'Pitch Limit System Check', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(88, 'Auto Pilot On and Check', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(89, 'Main Rotor Control speed Check range', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(90, 'Main Rotor RPM Set 95 %', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(91, 'Collective Down', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(92, 'Engine Startup checklist Completed', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(93, 'BEFORE TAXI', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(94, 'Taxing clearance ATC Request', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(95, 'Area Clear', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(96, 'Chocks   Removed and Stowed', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(97, 'Crew And Pax  Briefed', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(98, 'Doors and Windows Closed', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(99, 'Cargo Compartment  Secured', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(100, 'Autopilot Off', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(101, 'Before Taxing Checklist  Completed', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(102, 'LINEUP POSITION', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(103, 'Area Clear', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(104, 'Obstacles in TakeOff Direction Absent', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(105, 'Gyro Same readings', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(106, 'Heading Set', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(107, 'Autopilot ON', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(108, 'Type of TakeOFF Decide', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(109, 'Request for TakeOFF Performed', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(110, 'Stop watch Press', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(111, 'BEFORE TAKE OFF', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(112, 'Fuel Selector    Service', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(113, 'Fuel Pump Lights  Check Off', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(114, 'Transponder Alt       -', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(115, 'Autopilot On', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(116, 'Engine And Transmission Checked', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(117, 'Before Takeoff Checklist Completed', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(118, 'AFTER TAKE OFF', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(119, 'Main rotor 95 plus or minus 2 percent', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(120, 'DPDS Above 50 meters OFF', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(121, 'Fuel consumption Monitored', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(122, 'Monitor every 15 to 20 meters', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(123, 'After takeoff checklist Completed', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(124, 'IN FLIGHT CHECKS', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(125, 'Power setting As  per graph', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(126, 'Fuel quantity Check', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(127, 'Other Parameters Normal operation', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(128, 'Flight Following Call', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(129, 'PRE LANDING CHECKS', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(130, 'Landing clearance Request', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(131, 'Runway Condition Known', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(132, 'Autopilot Alt channel Off', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(133, 'Cargo Secured', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(134, 'Landing lights On', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(135, 'Fuel quantity Check', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(136, 'Compass Matched', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(137, 'dpds On At 50m', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(138, 'Type of Landing Decide', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(139, 'Runway Exit and Parking As instructed', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(140, 'Parameters Within Limit', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(141, 'Main rotor 95 plus or minus 2 percent', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(142, 'Landing Gear Brakes Released', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(143, 'dpds On', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(144, 'Cargo compartment Secured', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(145, 'Crew Briefed', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(146, 'Before landing checklist Completed', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(147, 'AFTER LANDING', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(148, 'Collective fully down', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(149, 'Auto pilot      Off', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(150, 'Parking as required', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(151, 'After landing checklist completed', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(152, 'ENGINE SHUT DOWN', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(153, 'Landing gear Brake Apply', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(154, 'Consumers Off', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(155, 'Rectifiers Off', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(156, 'Inverters Neutral', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(157, 'Generators Off', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(158, 'Throttle Fully Left', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(159, 'Left and Right Fuel Pumps Off', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(160, 'Engines Cooling 1 to 2 Minutes', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(161, 'Eng  Shutdown Levers Backward', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(162, 'Stopwatch Press', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(163, 'Ngg equal 0 at 35 Second Minimum', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(164, 'MAIN ROTOR STOP', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(165, 'At less or equal to 15% Main Rotor Slowly apply Main Rotor brake', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(166, 'At  Main Rotor stop  Controls move back and forth', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(167, 'Fire extinguisher  switch Off', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(168, 'Fuel shut off valves Leave open position', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(169, 'Service tank pump Off', '2025-12-26 13:18:17', '2025-12-26 13:18:17'),
(170, 'Batteries 1 and 2 Off', '2025-12-26 13:18:17', '2025-12-26 13:18:17');

-- --------------------------------------------------------

--
-- Table structure for table `starting_with_dc_gpu_checklist`
--

CREATE TABLE `starting_with_dc_gpu_checklist` (
  `id` int(11) NOT NULL,
  `name` varchar(100) NOT NULL,
  `created_at` timestamp NOT NULL DEFAULT current_timestamp(),
  `updated_at` timestamp NOT NULL DEFAULT current_timestamp() ON UPDATE current_timestamp()
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

--
-- Dumping data for table `starting_with_dc_gpu_checklist`
--

INSERT INTO `starting_with_dc_gpu_checklist` (`id`, `name`, `created_at`, `updated_at`) VALUES
(2, 'MI-17V-5 HELICOPTER START UP WITH DC GPU & INFLIGHT CHECKLIST', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(3, 'PRE APU START', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(4, 'Instruments and all switches As required', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(5, 'GPU Connected check and On', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(6, 'Circuit Breakers On', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(7, 'FDR and CVR On', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(8, 'Headsets Connected', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(9, 'Intercom Readability Check', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(10, 'ADF On', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(11, 'GPS On', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(12, 'Transponder On', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(13, 'ELT Arming', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(14, 'Global satellite tracking system On', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(15, 'TCAS On', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(16, 'Aircraft records on board and filled', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(17, 'Overhead hatch closed', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(18, 'Windscreens Clean', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(19, 'Seatbelts Fastened', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(20, 'Pneumatic System Check Pressure', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(21, 'Pedals Neutral', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(22, 'Cyclic Stick Neutral', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(23, 'Landing gear brakes applied', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(24, 'Collective pitch fully down', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(25, 'Throttle Twist Grip Fully Left', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(26, 'Friction Clutch Adjusted', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(27, 'Separate Throttle Lever Middle and Latched', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(28, 'Main Rotor Brake Fully Down', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(29, 'Engine Shut Down Levers Backward', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(30, 'Fuel Quantity as per mission', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(31, 'Warning Lights Check', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(32, 'Fire extinguishing System Check', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(33, 'Voice Warning System Check', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(34, 'Inverter On', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(35, 'EGT Indicators Test', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(36, 'Engine vibration system test', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(37, 'Inverter Off', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(38, 'Fire extinguishing System On', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(39, 'All pumps On', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(40, 'Generators Off', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(41, 'Fire Fuel shut off valve Check and  Open', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(42, 'Engine Shut Down Levers Backward', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(43, 'Pre APU Start Checklist Completed', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(44, 'APU START CHECKLIST', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(45, 'Startup clearance ATC Request', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(46, 'Ground crew Signal', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(47, 'APU Start', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(48, 'Stopwatch start', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(49, 'APU Parameters Check as Required', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(50, 'Battery 1 and 2,standby generator and equipment test On', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(51, 'Rectifiers On', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(52, 'GPU Disconnected', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(53, 'Startup checklist   Completed', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(54, 'ENGINE START UP', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(55, 'Start up Area Clear', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(56, 'Anti collision light  On', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(57, 'Engine selection as per wind direction', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(58, 'Ground crew  Signal', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(59, 'Engine Start button press', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(60, 'HP cock open', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(61, 'Stopwatch Set', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(62, 'Warm up engine 1 to 2 minutes', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(63, 'Engine Parameters as required', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(64, 'Start second engine as above', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(65, 'Warm up second engine 1 to 2 minutes', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(66, 'Idle parameters Check', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(67, 'DPD On', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(68, 'EGT Air Check', '2025-12-26 13:21:48', '2025-12-26 13:21:48'),
(69, 'FUNCTIONAL CHECK', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(70, 'Hydraulic system Check', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(71, 'Controls Response Check', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(72, 'EEG Test', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(73, 'Partial acceleration Check', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(74, 'ENGINE OPERATIONAL MODE', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(75, 'Throttle Twist Grip Fully Right', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(76, 'Generators Test and On', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(77, 'APU Generator Off', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(78, 'Inverters auto position', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(79, 'APU Off', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(80, 'Navigation Lights On', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(81, 'Blade tip lights Night and Poor visibility On', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(82, 'Cabin Lighting Night and Poor visibility On', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(83, 'Formation lights Night and Poor visibility On', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(84, 'Voice Warning System On', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(85, 'Gyro Horizons On', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(86, 'Compass System On and Check', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(87, 'Radio Altimeter On and Check', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(88, 'Baro Altimeter Set', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(89, 'Pitch Limit System Check', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(90, 'Auto Pilot On and Check', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(91, 'Main Rotor Control speed Check range', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(92, 'Main Rotor RPM Set 95 percent', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(93, 'Collective  Pitch Down', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(94, 'Engine Startup checklist Completed', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(95, 'BEFORE TAXI', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(96, 'Taxing clearance ATC Request', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(97, 'Area Clear', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(98, 'Chocks Removed and Stowed', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(99, 'Crew And Pax  Briefed', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(100, 'Doors and Windows Closed', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(101, 'Cargo Compartment  Secured', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(102, 'Autopilot Off', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(103, 'Before Taxing Checklist  Completed', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(104, 'LINEUP POSITION', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(105, 'Area Clear', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(106, 'Obstacles in TakeOff Direction Absent', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(107, 'Gyro Same readings', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(108, 'Heading Set', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(109, 'Autopilot ON', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(110, 'Type of TakeOFF Decide', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(111, 'Request for TakeOFF Performed', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(112, 'Stopwatch Press', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(113, 'BEFORE TAKEOFF', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(114, 'Fuel Selector    Service', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(115, 'Fuel Pump Lights  Check Off', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(116, 'Transponder     Altitude', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(117, 'Autopilot On', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(118, 'Engine And Transmission Checked', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(119, 'Before Takeoff Checklist Completed', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(120, 'AFTER TAKEOFF', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(121, 'Main rotor 95 plus or minus 2 percent', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(122, 'DPDS Above 50 meters OFF', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(123, 'Fuel consumption Monitored', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(124, 'Monitor every 15 to 20 minutes', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(125, 'After takeoff checklist Completed', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(126, 'IN FLIGHT CHECKS', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(127, 'Power setting As  per graph', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(128, 'Fuel quantity Check', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(129, 'Other Parameters Normal operation', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(130, 'Flight Following Call', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(131, 'PRE LANDING CHECKS', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(132, 'Landing clearance Request', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(133, 'Runway Condition Known', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(134, 'Autopilot Alt channel Off', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(135, 'Cargo Secured', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(136, 'Landing lights On', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(137, 'Fuel quantity Check', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(138, 'Compass Matched', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(139, 'DPDS On At 50 meters', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(140, 'Type of Landing Decide', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(141, 'Runway Exit and Parking As instructed', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(142, 'Parameters Within Limit', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(143, 'Main rotor 95 plus or minus 2 percent', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(144, 'Landing Gear Brakes Released', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(145, 'DPDS On', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(146, 'Cargo compartment Secured', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(147, 'Crew Briefed', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(148, 'Before landing checklist Completed', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(149, 'AFTER LANDING', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(150, 'Collective FULLY DOWN', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(151, 'Auto pilot      OFF', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(152, 'Parking AS REQURED', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(153, 'After landing checklist COMPLETED', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(154, 'ENGINE SHUT DOWN', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(155, 'Landing gear Brake Apply', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(156, 'Consumers Off', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(157, 'Rectifiers Off', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(158, 'Inverters Neutral', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(159, 'Generators Off', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(160, 'Throttle Fully Left', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(161, 'Left and Right Fuel Pumps Off', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(162, 'Engines Cooling 1 to 2 Minutes', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(163, 'Engine  Shutdown Levers Backward', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(164, 'Stopwatch Press', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(165, 'Ngg equal 0 at 35 Second Minimum', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(166, 'MAIN ROTOR STOP', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(167, 'At less or equal to 15 % Main Rotor Slowly apply Main Rotor brake', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(168, 'At Main Rotor stop Controls move back and forth', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(169, 'Fire extinguisher switch Off', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(170, 'Fuel shutoff valves Leave open position', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(171, 'Service tank pump Off', '2025-12-26 13:21:49', '2025-12-26 13:21:49'),
(172, 'Batteries 1 and 2 Off', '2025-12-26 13:21:49', '2025-12-26 13:21:49');

-- --------------------------------------------------------

--
-- Table structure for table `status`
--

CREATE TABLE `status` (
  `id` int(11) NOT NULL,
  `name` varchar(20) NOT NULL,
  `created_at` timestamp NOT NULL DEFAULT current_timestamp(),
  `updated_at` timestamp NOT NULL DEFAULT current_timestamp() ON UPDATE current_timestamp()
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

--
-- Dumping data for table `status`
--

INSERT INTO `status` (`id`, `name`, `created_at`, `updated_at`) VALUES
(0, 'inactive', '2025-12-26 12:51:47', '2025-12-26 12:55:32'),
(1, 'active', '2025-12-26 12:51:47', '2025-12-26 12:51:47');

--
-- Indexes for dumped tables
--

--
-- Indexes for table `aircrafts`
--
ALTER TABLE `aircrafts`
  ADD PRIMARY KEY (`id`),
  ADD UNIQUE KEY `call_sign` (`call_sign`),
  ADD KEY `idx_aircrafts_type` (`aircraft_type_id`),
  ADD KEY `idx_aircrafts_status` (`status_id`);

--
-- Indexes for table `aircraft_categories`
--
ALTER TABLE `aircraft_categories`
  ADD PRIMARY KEY (`id`),
  ADD UNIQUE KEY `name` (`name`);

--
-- Indexes for table `aircraft_types`
--
ALTER TABLE `aircraft_types`
  ADD PRIMARY KEY (`id`),
  ADD UNIQUE KEY `type` (`type`),
  ADD KEY `aircraft_category_id` (`aircraft_category_id`);

--
-- Indexes for table `anomalies`
--
ALTER TABLE `anomalies`
  ADD PRIMARY KEY (`id`),
  ADD KEY `phase_of_flight_id` (`phase_of_flight_id`),
  ADD KEY `idx_anomalies_flight` (`flight_id`),
  ADD KEY `fk_parameter_MI_17V_5_name` (`parameter_MI_17V_5_name`);

--
-- Indexes for table `checklist_types`
--
ALTER TABLE `checklist_types`
  ADD PRIMARY KEY (`id`),
  ADD UNIQUE KEY `name` (`name`);

--
-- Indexes for table `crews`
--
ALTER TABLE `crews`
  ADD PRIMARY KEY (`id`),
  ADD UNIQUE KEY `code` (`code`),
  ADD KEY `idx_crews_type` (`crew_type_id`),
  ADD KEY `fk_status_id` (`status_id`);

--
-- Indexes for table `crew_types`
--
ALTER TABLE `crew_types`
  ADD PRIMARY KEY (`id`),
  ADD UNIQUE KEY `name` (`name`);

--
-- Indexes for table `exceedances`
--
ALTER TABLE `exceedances`
  ADD PRIMARY KEY (`id`),
  ADD KEY `idx_exceedances_flight` (`flight_id`),
  ADD KEY `fk_parameter_MI_17V_5_name_real` (`parameter_MI_17V_5_name`);

--
-- Indexes for table `flights`
--
ALTER TABLE `flights`
  ADD PRIMARY KEY (`id`),
  ADD KEY `PIC` (`PIC`),
  ADD KEY `SIC` (`SIC`),
  ADD KEY `FE` (`FE`),
  ADD KEY `flight_type_id` (`flight_type_id`),
  ADD KEY `idx_flights_aircraft` (`aircraft_id`),
  ADD KEY `idx_flights_date` (`flight_date`);

--
-- Indexes for table `flight_types`
--
ALTER TABLE `flight_types`
  ADD PRIMARY KEY (`id`),
  ADD UNIQUE KEY `name` (`name`);

--
-- Indexes for table `missed_checks`
--
ALTER TABLE `missed_checks`
  ADD PRIMARY KEY (`id`),
  ADD KEY `checklist_type_id` (`checklist_type_id`),
  ADD KEY `idx_missed_checks_flight` (`flight_id`);

--
-- Indexes for table `parameters`
--
ALTER TABLE `parameters`
  ADD PRIMARY KEY (`id`),
  ADD UNIQUE KEY `MI_17V_5_name` (`MI_17V_5_name`),
  ADD UNIQUE KEY `MI_17_1V_name` (`MI_17_1V_name`),
  ADD KEY `idx_parameters_aircraft_type` (`aircraft_type_id`);

--
-- Indexes for table `phase_of_flights`
--
ALTER TABLE `phase_of_flights`
  ADD PRIMARY KEY (`id`),
  ADD UNIQUE KEY `name` (`name`);

--
-- Indexes for table `starting_without_gpu_checklist`
--
ALTER TABLE `starting_without_gpu_checklist`
  ADD PRIMARY KEY (`id`);

--
-- Indexes for table `starting_with_ac_gpu_checklist`
--
ALTER TABLE `starting_with_ac_gpu_checklist`
  ADD PRIMARY KEY (`id`);

--
-- Indexes for table `starting_with_dc_gpu_checklist`
--
ALTER TABLE `starting_with_dc_gpu_checklist`
  ADD PRIMARY KEY (`id`);

--
-- Indexes for table `status`
--
ALTER TABLE `status`
  ADD PRIMARY KEY (`id`),
  ADD UNIQUE KEY `name` (`name`);

--
-- AUTO_INCREMENT for dumped tables
--

--
-- AUTO_INCREMENT for table `aircrafts`
--
ALTER TABLE `aircrafts`
  MODIFY `id` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=7;

--
-- AUTO_INCREMENT for table `aircraft_categories`
--
ALTER TABLE `aircraft_categories`
  MODIFY `id` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=3;

--
-- AUTO_INCREMENT for table `aircraft_types`
--
ALTER TABLE `aircraft_types`
  MODIFY `id` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=3;

--
-- AUTO_INCREMENT for table `anomalies`
--
ALTER TABLE `anomalies`
  MODIFY `id` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=41;

--
-- AUTO_INCREMENT for table `checklist_types`
--
ALTER TABLE `checklist_types`
  MODIFY `id` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=4;

--
-- AUTO_INCREMENT for table `crews`
--
ALTER TABLE `crews`
  MODIFY `id` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=47;

--
-- AUTO_INCREMENT for table `crew_types`
--
ALTER TABLE `crew_types`
  MODIFY `id` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=4;

--
-- AUTO_INCREMENT for table `exceedances`
--
ALTER TABLE `exceedances`
  MODIFY `id` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=41;

--
-- AUTO_INCREMENT for table `flights`
--
ALTER TABLE `flights`
  MODIFY `id` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=12;

--
-- AUTO_INCREMENT for table `flight_types`
--
ALTER TABLE `flight_types`
  MODIFY `id` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=4;

--
-- AUTO_INCREMENT for table `missed_checks`
--
ALTER TABLE `missed_checks`
  MODIFY `id` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=211;

--
-- AUTO_INCREMENT for table `parameters`
--
ALTER TABLE `parameters`
  MODIFY `id` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=100;

--
-- AUTO_INCREMENT for table `phase_of_flights`
--
ALTER TABLE `phase_of_flights`
  MODIFY `id` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=4;

--
-- AUTO_INCREMENT for table `starting_without_gpu_checklist`
--
ALTER TABLE `starting_without_gpu_checklist`
  MODIFY `id` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=171;

--
-- AUTO_INCREMENT for table `starting_with_ac_gpu_checklist`
--
ALTER TABLE `starting_with_ac_gpu_checklist`
  MODIFY `id` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=171;

--
-- AUTO_INCREMENT for table `starting_with_dc_gpu_checklist`
--
ALTER TABLE `starting_with_dc_gpu_checklist`
  MODIFY `id` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=173;

--
-- AUTO_INCREMENT for table `status`
--
ALTER TABLE `status`
  MODIFY `id` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=3;

--
-- Constraints for dumped tables
--

--
-- Constraints for table `aircrafts`
--
ALTER TABLE `aircrafts`
  ADD CONSTRAINT `aircrafts_ibfk_1` FOREIGN KEY (`aircraft_type_id`) REFERENCES `aircraft_types` (`id`),
  ADD CONSTRAINT `aircrafts_ibfk_2` FOREIGN KEY (`status_id`) REFERENCES `status` (`id`);

--
-- Constraints for table `aircraft_types`
--
ALTER TABLE `aircraft_types`
  ADD CONSTRAINT `aircraft_types_ibfk_1` FOREIGN KEY (`aircraft_category_id`) REFERENCES `aircraft_categories` (`id`);

--
-- Constraints for table `anomalies`
--
ALTER TABLE `anomalies`
  ADD CONSTRAINT `anomalies_ibfk_1` FOREIGN KEY (`flight_id`) REFERENCES `flights` (`id`),
  ADD CONSTRAINT `anomalies_ibfk_3` FOREIGN KEY (`phase_of_flight_id`) REFERENCES `phase_of_flights` (`id`),
  ADD CONSTRAINT `fk_parameter_MI_17V_5_name` FOREIGN KEY (`parameter_MI_17V_5_name`) REFERENCES `parameters` (`MI_17V_5_name`) ON DELETE CASCADE ON UPDATE CASCADE;

--
-- Constraints for table `crews`
--
ALTER TABLE `crews`
  ADD CONSTRAINT `crews_ibfk_1` FOREIGN KEY (`crew_type_id`) REFERENCES `crew_types` (`id`),
  ADD CONSTRAINT `fk_status_id` FOREIGN KEY (`status_id`) REFERENCES `status` (`id`) ON DELETE CASCADE ON UPDATE CASCADE;

--
-- Constraints for table `exceedances`
--
ALTER TABLE `exceedances`
  ADD CONSTRAINT `exceedances_ibfk_1` FOREIGN KEY (`flight_id`) REFERENCES `flights` (`id`),
  ADD CONSTRAINT `fk_parameter_MI_17V_5_name_real` FOREIGN KEY (`parameter_MI_17V_5_name`) REFERENCES `parameters` (`MI_17V_5_name`) ON DELETE CASCADE ON UPDATE CASCADE;

--
-- Constraints for table `flights`
--
ALTER TABLE `flights`
  ADD CONSTRAINT `flights_ibfk_1` FOREIGN KEY (`aircraft_id`) REFERENCES `aircrafts` (`id`),
  ADD CONSTRAINT `flights_ibfk_2` FOREIGN KEY (`PIC`) REFERENCES `crews` (`code`),
  ADD CONSTRAINT `flights_ibfk_3` FOREIGN KEY (`SIC`) REFERENCES `crews` (`code`),
  ADD CONSTRAINT `flights_ibfk_4` FOREIGN KEY (`FE`) REFERENCES `crews` (`code`),
  ADD CONSTRAINT `flights_ibfk_5` FOREIGN KEY (`flight_type_id`) REFERENCES `flight_types` (`id`);

--
-- Constraints for table `missed_checks`
--
ALTER TABLE `missed_checks`
  ADD CONSTRAINT `missed_checks_ibfk_1` FOREIGN KEY (`flight_id`) REFERENCES `flights` (`id`),
  ADD CONSTRAINT `missed_checks_ibfk_2` FOREIGN KEY (`checklist_type_id`) REFERENCES `checklist_types` (`id`);

--
-- Constraints for table `parameters`
--
ALTER TABLE `parameters`
  ADD CONSTRAINT `parameters_ibfk_1` FOREIGN KEY (`aircraft_type_id`) REFERENCES `aircraft_types` (`id`);
COMMIT;

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
