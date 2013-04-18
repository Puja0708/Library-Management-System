--
-- Database: `Library`
--

-- --------------------------------------------------------

--
-- Table structure for table `Issue_return`
--

CREATE TABLE IF NOT EXISTS issue_return(
  Bno int(15) NOT NULL,
  ID int(15) NOT NULL,
  Issue_date date NOT NULL,
  Due_date date NOT NULL,
  Return_date date NOT NULL,
  fine int(3) NOT NULL,
  copies_number int(3) NOT NULL,
  PRIMARY KEY (ID),
  UNIQUE KEY Bno (Bno)
);

--
-- Dumping data for table `Issue_return`
--

